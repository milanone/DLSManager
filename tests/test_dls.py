"""Tests for the DLS analysis of the stored correlogram: synthetic correlograms of known sizes (no files, no Tk),
and the real sample in 'example data/' against the values of the instrument (skipped if the file is missing).

    py -m unittest discover -s tests -v
"""
import importlib.util
import os
import sys
import tkinter as tk
import unittest

import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("nsz", os.path.join(HERE, "..", "DLSManager.pyw"))
nsz = importlib.util.module_from_spec(spec)
spec.loader.exec_module(nsz)

sys.path.insert(0, HERE)
from test_nsz_viewer import SAMPLES      # (.nsz, CSV) pairs of 'example data/'


try:
    import scipy.optimize  # noqa: F401
    HAS_SCIPY = True
except ImportError:
    HAS_SCIPY = False

# the delays of the SZ-100 correlator (seconds) and its usual conditions
TAU = np.array([0.1, 0.2, 0.3, 0.4, 0.5, 0.7, 1.0, 1.4, 1.9, 2.6, 3.6, 5.0, 6.9, 9.5, 13.1, 18.1, 25.0, 34.5, 47.6, 65.7,
                90.7, 125.2, 172.8, 238.5, 329.2, 454.4, 627.2, 865.7, 1194.9, 1649.3, 2276.5, 3142.2, 4337.1, 5986.4,
                8262.9, 11405.1, 15742.2, 21728.6, 29991.5, 41396.6, 57138.8, 78867.4, 108858.9]) * 1e-6
COND = {'angle': 173.0, 'n': 1.333, 'visc': 0.8963, 'T': 25.0}


def correlogram(sizes, fractions, intercept=0.99):
    """Normalised g2-1 of a mixture of spheres (diameters in nm, intensity fractions)."""
    q2 = nsz.dls_q2(COND['angle'], COND['n'], 532.0)
    gam = nsz._dls_gamma(sizes, q2, COND['T'], COND['visc'])
    g1 = intercept * (np.exp(-np.outer(TAU, gam)) @ np.asarray(fractions, float))
    return g1 ** 2


def record(sizes, fractions):
    return dict(COND, tau=TAU, acf=correlogram(sizes, fractions))


class TestCumulants(unittest.TestCase):
    def test_monodisperse_gives_its_size_and_no_polydispersity(self):
        q2 = nsz.dls_q2(COND['angle'], COND['n'], 532.0)
        z, pi = nsz.dls_cumulants(TAU, correlogram([120.0], [1.0]), q2, COND['T'], COND['visc'])
        self.assertAlmostEqual(z, 120.0, delta=3.0)
        self.assertLess(pi, 0.02)

    def test_too_few_points_gives_none(self):
        q2 = nsz.dls_q2(173.0, 1.333)
        self.assertIsNone(nsz.dls_cumulants(TAU, np.full(len(TAU), 0.01), q2, 25.0, 0.89))

    def test_q2_scales_with_the_inverse_square_of_the_wavelength(self):
        self.assertAlmostEqual(nsz.dls_q2(90, 1.333, 532) / nsz.dls_q2(90, 1.333, 1064), 4.0, places=6)


@unittest.skipUnless(HAS_SCIPY, "scipy not installed")
class TestDistributions(unittest.TestCase):
    def test_the_free_model_fits_the_intercept_of_g1(self):
        # a correlogram that starts at 0.9, not 1: the fit has to start there too (it was forced to 1)
        rec = record([100.0], [1.0])
        rec['acf'] = rec['acf'] * 0.9 ** 2 / 0.99 ** 2
        r = nsz.dls_analyse(rec, nsz.DLS_MODELS[0])
        used = rec['tau'] >= 1e-6
        first = np.nonzero(used)[0][0]
        self.assertAlmostEqual(r['fit'][first], 0.9 * np.sqrt(1.0) * (np.exp(-1e-6 * nsz._dls_gamma(100.0, nsz.dls_q2(173.0, 1.333), 25.0, 0.8963))), delta=0.02)
        self.assertLess(r['fit'][first], 0.95)
        self.assertAlmostEqual(r['p'].sum(), 1.0)             # the distribution itself stays normalised

    def test_regularised_peak_is_at_the_true_size(self):
        r = nsz.dls_analyse(record([100.0], [1.0]), nsz.DLS_MODELS[0])
        self.assertAlmostEqual(r['p'].sum(), 1.0, places=6)
        self.assertAlmostEqual(r['mode'], 100.0, delta=25.0)
        self.assertIsNone(r['peaks'])

    def test_one_peak_model(self):
        r = nsz.dls_analyse(record([100.0], [1.0]), nsz.DLS_MODELS[1])
        self.assertEqual(len(r['peaks']), 1)
        self.assertAlmostEqual(r['peaks'][0]['median'], 100.0, delta=8.0)
        self.assertAlmostEqual(r['peaks'][0]['fraction'], 1.0)

    def test_free_intercept_removes_the_spurious_small_peak(self):
        # The stored ACF starts near 0.99, not 1: without a free intercept a 3-5% peak appeared at a few nm.
        r = nsz.dls_analyse(record([100.0], [1.0]), nsz.DLS_MODELS[2])
        small = [pk for pk in r['peaks'] if pk['median'] < 30.0 and pk['fraction'] > 0.02]
        self.assertEqual(small, [])

    def test_two_peaks_well_separated(self):
        rec = record([60.0, 400.0], [0.5, 0.5])
        two = nsz.dls_analyse(rec, nsz.DLS_MODELS[2])
        one = nsz.dls_analyse(rec, nsz.DLS_MODELS[1])
        self.assertEqual(len(two['peaks']), 2)
        self.assertAlmostEqual(two['peaks'][0]['median'], 60.0, delta=20.0)
        self.assertAlmostEqual(two['peaks'][1]['median'], 400.0, delta=100.0)
        self.assertLess(two['bic'], one['bic'] - 6)            # two peaks are clearly supported

    def test_one_peak_is_enough_for_one_size(self):
        rec = record([150.0], [1.0])
        two = nsz.dls_analyse(rec, nsz.DLS_MODELS[2])
        one = nsz.dls_analyse(rec, nsz.DLS_MODELS[1])
        self.assertGreater(two['bic'], one['bic'] - 6)         # the second peak is not significant


class TestRecord(unittest.TestCase):
    def test_unaligned_or_incomplete_samples_give_none(self):
        d = {'tau': TAU * 1e6, 'acf': np.ones(len(TAU)), 'tau_aligned': False,
             'params': {'Angle_deg': 173.0, 'n_solvent': 1.333, 'Visc_mPas': 0.9, 'Temp_C': 25.0}}
        self.assertIsNone(nsz.dls_record(d))
        d['tau_aligned'] = True
        self.assertAlmostEqual(nsz.dls_record(d)['tau'][6], 1e-6)
        d['params']['Visc_mPas'] = None
        self.assertIsNone(nsz.dls_record(d))


@unittest.skipUnless(SAMPLES and HAS_SCIPY, "real samples or scipy not present")
class TestRealSample(unittest.TestCase):
    """Every sample pair of 'example data/' (a polystyrene latex and two liposome preparations, two software layouts)."""

    def each(self):
        for path, _ in SAMPLES:
            with self.subTest(sample=os.path.basename(path)):
                d = nsz.load_nsz(path)
                yield d, nsz.dls_record(d)

    def test_record_uses_the_conditions_of_the_file(self):
        for d, rec in self.each():
            self.assertEqual(rec['angle'], d['params']['Angle_deg'])
            self.assertEqual(len(rec['tau']), 43)

    def test_z_average_is_close_to_the_instrument(self):
        # 1-2% for the narrow samples, 4% at most for the polydisperse liposomes and the 1:400 latex
        for d, rec in self.each():
            r = nsz.dls_analyse(rec, nsz.DLS_MODELS[0], d['params']['Lambda_nm'])
            self.assertAlmostEqual(r['z'] / d['params']['ZAvg_nm'], 1.0, delta=0.05)

    def test_all_models_run(self):
        for d, rec in self.each():
            for m in nsz.DLS_MODELS:
                r = nsz.dls_analyse(rec, m, 532.0)
                self.assertTrue(np.all(np.isfinite(r['p'])))
                self.assertEqual(len(r['fit']), 43)


# Francesco moved the DLS samples on 2026-11-23 from 'spettri di esempio' to 'size da spettro scattering' (LabSpectrumManager)
_LSM = os.path.join(HERE, "..", "..", "LabSpectrumManager")
_ROOT = next((os.path.join(_LSM, d) for d in ("size da spettro scattering", "spettri di esempio")
              if os.path.isdir(os.path.join(_LSM, d, "20261119 DLS lattice e liposomi"))), os.path.join(_LSM, "spettri di esempio"))
LATEX_DIR = os.path.join(_ROOT, "20261119 DLS lattice e liposomi")
LATEX_173 = [os.path.join(LATEX_DIR, "Polistirene 1a100 in NaCl 10 mM_084%d.nsz" % i) for i in (0, 1, 2)]
HAS_LATEX = all(os.path.isfile(f) for f in LATEX_173)


@unittest.skipUnless(HAS_LATEX and HAS_SCIPY, "polystyrene latex files (LabSpectrumManager sample data) not present")
class TestLatex(unittest.TestCase):
    """Polystyrene latex (certified 130 nm by TEM), 1:100 in 10 mM NaCl, 173 degrees: a narrow, single-size sample.
    These files keep the measurement parameters in Base35 (other software version), not Base34."""

    @classmethod
    def setUpClass(cls):
        cls.data = [nsz.load_nsz(f) for f in LATEX_173]

    def test_parameters_are_found_in_the_newer_stream(self):
        for d in self.data:
            p = d['params']
            self.assertEqual(p['Angle_deg'], 173.0)
            self.assertAlmostEqual(p['n_solvent'], 1.333, places=3)
            self.assertAlmostEqual(p['Temp_C'], 25.0, delta=0.3)
            self.assertEqual(p['Lambda_nm'], 532.0)
            self.assertTrue(p['Solvent'])
            self.assertTrue(d['tau_aligned'])

    def test_z_average_matches_the_instrument_within_one_percent(self):
        for d in self.data:
            r = nsz.dls_analyse(nsz.dls_record(d), nsz.DLS_MODELS[0], d['params']['Lambda_nm'])
            self.assertAlmostEqual(r['z'] / d['params']['ZAvg_nm'], 1.0, delta=0.01)

    def test_distribution_is_one_narrow_peak_at_the_instrument_size(self):
        for d in self.data:
            rec = nsz.dls_record(d)
            reg = nsz.dls_analyse(rec, nsz.DLS_MODELS[0], 532.0)
            self.assertAlmostEqual(reg['mode'] / d['params']['Mode_nm'], 1.0, delta=0.05)
            one = nsz.dls_analyse(rec, nsz.DLS_MODELS[1], 532.0)
            two = nsz.dls_analyse(rec, nsz.DLS_MODELS[2], 532.0)
            self.assertAlmostEqual(one['peaks'][0]['median'] / d['params']['Mode_nm'], 1.0, delta=0.05)
            self.assertLess(one['peaks'][0]['sigma'], 0.35)
            self.assertLess(one['bic'] - two['bic'], 6)        # a second peak is not supported


class TestMie(unittest.TestCase):
    def test_small_particles_follow_the_d6_law(self):
        d = np.array([5.0, 10.0, 20.0])
        i = nsz.mie_intensity(d, 532.0, 1.333, 1.59, 90.0)
        ratio = i / d ** 6
        np.testing.assert_allclose(ratio, ratio[0], rtol=0.01)

    def test_matches_an_independent_sum_of_the_coefficients(self):
        # Forward scattering: Qext = 4/x^2 Re S(0), with S(0) = sum (2n+1)(a+b)/2, and Qext = Qsca without absorption.
        # Rebuilt here from scipy's spherical Bessel functions, independently of the recurrences in the program.
        from scipy.special import spherical_jn, spherical_yn
        n_med, n_p, wl, d = 1.333, 1.59, 532.0, 400.0
        x, m = np.pi * d * n_med / wl, n_p / n_med
        n = np.arange(1, int(x + 4 * x ** (1 / 3) + 2) + 1)

        def psi(z, k):
            return z * spherical_jn(k, z)

        def dpsi(z, k):
            return spherical_jn(k, z) + z * spherical_jn(k, z, derivative=True)

        def xi(z, k):
            return z * (spherical_jn(k, z) + 1j * spherical_yn(k, z))

        def dxi(z, k):
            h = spherical_jn(k, z) + 1j * spherical_yn(k, z)
            return h + z * (spherical_jn(k, z, derivative=True) + 1j * spherical_yn(k, z, derivative=True))

        a = (m * psi(m * x, n) * dpsi(x, n) - psi(x, n) * dpsi(m * x, n)) / (m * psi(m * x, n) * dxi(x, n) - xi(x, n) * dpsi(m * x, n))
        b = (psi(m * x, n) * dpsi(x, n) - m * psi(x, n) * dpsi(m * x, n)) / (psi(m * x, n) * dxi(x, n) - m * xi(x, n) * dpsi(m * x, n))
        s0 = np.sum((2 * n + 1) * (a + b)) / 2
        qext, qsca = 4 / x ** 2 * s0.real, 2 / x ** 2 * np.sum((2 * n + 1) * (abs(a) ** 2 + abs(b) ** 2))
        self.assertAlmostEqual(qext, qsca, places=6)                    # no absorption
        i0 = nsz.mie_intensity(np.array([d]), wl, n_med, n_p, 0.0)[0]   # the program's S1(0)
        self.assertAlmostEqual(i0, abs(s0) ** 2, delta=1e-6 * abs(s0) ** 2)

    def test_mie_conversion_differs_from_rayleigh_for_large_particles(self):
        d = np.geomspace(20, 2000, 60)
        i1 = nsz.mie_intensity(d, 532.0, 1.333, 1.59, 173.0)
        self.assertTrue(np.all(i1 > 0))
        p = np.exp(-0.5 * ((np.log(d) - np.log(300)) / 0.3) ** 2)
        p /= p.sum()
        mie = nsz.dls_stats(d, nsz.dls_reweight(d, p, "Number", i1))[1]
        ray = nsz.dls_stats(d, nsz.dls_reweight(d, p, "Number"))[1]
        self.assertNotAlmostEqual(mie, ray, delta=5.0)
        # a per-particle intensity of d^6 reproduces the Rayleigh result exactly, peaks included
        np.testing.assert_allclose(nsz.dls_reweight(d, p, "Volume", d ** 6), nsz.dls_reweight(d, p, "Volume"))
        peaks = [{'median': 300.0, 'sigma': 0.3, 'fraction': 1.0}]
        a = nsz.dls_reweight_peaks(peaks, "Number")[0]['median']
        b = nsz.dls_reweight_peaks(peaks, "Number", np.geomspace(5, 1e4, 4000), np.geomspace(5, 1e4, 4000) ** 6)[0]['median']
        self.assertAlmostEqual(a, b, delta=0.01 * a)


class TestAperture(unittest.TestCase):
    REC = {'angle': 90.0, 'n': 1.333}

    def test_zero_aperture_is_the_single_angle(self):
        d = np.geomspace(60, 500, 12)
        a = nsz.dls_per_particle(d, "Mie", self.REC, 532.0, 1.59, aperture=0.0)
        np.testing.assert_allclose(a, nsz.mie_intensity(d, 532.0, 1.333, 1.59, 90.0))

    def test_aperture_fills_the_interference_minimum_and_leaves_rayleigh_alone(self):
        # latex at 90 degrees: one angle scatters 5 times less at 350 nm than at 250 nm, which turned the small tail of
        # a free distribution at 350 nm into a spurious volume peak; the aperture average removes that dip
        d = np.array([250.0, 350.0])
        single = nsz.dls_per_particle(d, "Mie", self.REC, 532.0, 1.59, aperture=0.0)
        wide = nsz.dls_per_particle(d, "Mie", self.REC, 532.0, 1.59, aperture=10.0)
        self.assertLess(single[1] / single[0], 0.3)
        self.assertGreater(wide[1] / wide[0], 1.5 * single[1] / single[0])      # 0.17 -> 0.33
        np.testing.assert_allclose(nsz.dls_per_particle(d, "Rayleigh", self.REC, 532.0, 1.59, aperture=10.0), d ** 6)
        # small particles: the angle does not matter for Rayleigh scatterers, so the average equals the single angle
        s = np.array([10.0, 20.0])
        np.testing.assert_allclose(nsz.dls_per_particle(s, "Mie", self.REC, 532.0, 1.59, aperture=10.0),
                                   nsz.dls_per_particle(s, "Mie", self.REC, 532.0, 1.59, aperture=0.0), rtol=0.02)


@unittest.skipUnless(HAS_SCIPY, "scipy not installed")
class TestMieShell(unittest.TestCase):
    def test_a_shell_as_thick_as_the_radius_is_a_solid_sphere(self):
        d = np.array([30.0, 130.0, 600.0])
        for angle in (90.0, 173.0):
            shell = nsz.mie_shell_intensity(d, 532.0, 1.333, 0.12, 1e5, angle)
            solid = nsz.mie_intensity(d, 532.0, 1.333, 1.333 + 0.12, angle)
            np.testing.assert_allclose(shell, solid, rtol=1e-6)

    def test_a_thin_shell_scatters_like_d4_not_d6(self):
        d = np.array([30.0, 60.0])
        i = nsz.mie_shell_intensity(d, 532.0, 1.333, 0.12, 4.5, 90.0)
        self.assertAlmostEqual(np.log(i[1] / i[0]) / np.log(2.0), 4.0, delta=0.5)

    def test_shell_conversion_is_available_in_the_tab_conversions(self):
        self.assertIn("Mie shell", nsz.DLS_CONVERSIONS)
        rec = dict(COND)
        d = np.geomspace(20, 800, 12)
        i1 = nsz.dls_per_particle(d, "Mie shell", rec, 532.0, 1.6, (4.5, 0.12))
        self.assertEqual(i1.shape, d.shape)
        self.assertTrue(np.all(i1 > 0))


class TestWeighting(unittest.TestCase):
    def test_reweight_matches_rayleigh_and_normalises(self):
        d = np.array([10.0, 100.0])
        p = np.array([0.5, 0.5])
        vol = nsz.dls_reweight(d, p, "Volume")
        num = nsz.dls_reweight(d, p, "Number")
        self.assertAlmostEqual(vol.sum(), 1.0)
        self.assertAlmostEqual(vol[0] / vol[1], 1000.0, places=6)         # (100/10)^3
        self.assertAlmostEqual(num[0] / num[1], 1e6, places=3)            # (100/10)^6
        np.testing.assert_allclose(nsz.dls_reweight(d, p, "Intensity"), p)

    def test_peaks_shift_by_exp_k_sigma_squared_and_agree_with_the_grid(self):
        peaks = [{'median': 60.0, 'sigma': 0.2, 'fraction': 0.3}, {'median': 400.0, 'sigma': 0.15, 'fraction': 0.7}]
        for kind, k in (("Volume", 3), ("Number", 6)):
            new = nsz.dls_reweight_peaks(peaks, kind)
            self.assertAlmostEqual(new[0]['median'], 60.0 * np.exp(-k * 0.04), places=6)
            self.assertAlmostEqual(sum(pk['fraction'] for pk in new), 1.0)
            # the same on a grid: share of the small peak after dividing the intensity by d^k
            d = np.geomspace(5, 1e4, 4000)
            lnd = np.log(d)
            g = [pk['fraction'] * np.exp(-0.5 * ((lnd - np.log(pk['median'])) / pk['sigma']) ** 2) /
                 (pk['sigma'] * np.sqrt(2 * np.pi)) for pk in peaks]
            w = [gi / d ** k for gi in g]
            self.assertAlmostEqual(w[0].sum() / (w[0].sum() + w[1].sum()), new[0]['fraction'], places=3)

    def test_number_weighting_moves_a_broad_distribution_to_smaller_sizes(self):
        d = np.geomspace(10, 3000, 100)
        p = np.exp(-0.5 * ((np.log(d) - np.log(300)) / 0.5) ** 2)
        p /= p.sum()
        mean_i, mode_i, _ = nsz.dls_stats(d, p)
        mean_v, mode_v, _ = nsz.dls_stats(d, nsz.dls_reweight(d, p, "Volume"))
        mean_n, mode_n, _ = nsz.dls_stats(d, nsz.dls_reweight(d, p, "Number"))
        self.assertGreater(mode_i, mode_v)
        self.assertGreater(mode_v, mode_n)


class TestWidth(unittest.TestCase):
    D = np.geomspace(10, 3000, 400)

    def _lognormal(self, median, sigma):
        p = np.exp(-0.5 * ((np.log(self.D) - np.log(median)) / sigma) ** 2)
        return p / p.sum()

    def _sigma(self, p):
        lnd = np.log(self.D)
        return float(np.sqrt(np.sum(p * (lnd - np.sum(p * lnd)) ** 2)))

    def test_set_width_gives_the_requested_sigma_to_narrower_and_wider_distributions(self):
        for sigma_in in (0.10, 0.45, 0.72):
            for target in (0.24, 0.43):
                p = self._lognormal(200.0, sigma_in)
                q = nsz.dls_set_width(self.D, p, target)
                self.assertAlmostEqual(q.sum(), 1.0)
                self.assertAlmostEqual(self._sigma(q), target, delta=0.012)
                self.assertAlmostEqual(float(np.exp(np.sum(q * np.log(self.D)))), 200.0, delta=3.0)

    def test_resolution_map_reproduces_the_measured_widths_and_decreases(self):
        # the widths read from the instrument's recalculated size tables (sigma of ln d) at the presets of its manual
        for res, sigma in ((-0.5, 0.43), (0.1, 0.30), (0.3, 0.24), (0.6, 0.07)):
            self.assertAlmostEqual(nsz.dls_resolution_sigma(res), sigma, places=6)
        values = [nsz.dls_resolution_sigma(r) for r in np.linspace(-1, 1, 41)]
        self.assertTrue(all(b <= a + 1e-12 for a, b in zip(values, values[1:])))
        self.assertTrue(all(0.03 <= v <= 1.0 for v in values))
        self.assertEqual({k: nsz.dls_resolution_sigma(v) for k, v in nsz.DLS_FORM_RESOLUTION.items()},
                         {"Broad": 0.43, "Standard": 0.24, "Narrow": 0.07})

    def test_number_mean_of_a_narrow_sample_recovered_by_giving_the_distribution_its_true_width(self):
        # the true number distribution is narrow (sigma 0.10); the regularised inversion gives sigma 0.31 around the same
        # intensity median. Number mean ~ median x exp(-6 sigma^2): 130 with the true width, ~73% of it with the broad one.
        true = self._lognormal(137.0, 0.10)
        broad = self._lognormal(137.0, 0.31)
        n_true = nsz.dls_stats(self.D, nsz.dls_reweight(self.D, true, "Number"))[0]
        n_broad = nsz.dls_stats(self.D, nsz.dls_reweight(self.D, broad, "Number"))[0]
        n_set = nsz.dls_stats(self.D, nsz.dls_reweight(self.D, nsz.dls_set_width(self.D, broad, 0.10), "Number"))[0]
        self.assertLess(n_broad, 0.75 * n_true)
        self.assertAlmostEqual(n_set, n_true, delta=0.03 * n_true)


@unittest.skipUnless(HAS_SCIPY, "scipy not installed")
class TestInstrumentRange(unittest.TestCase):
    def test_range_limits_the_points_the_inversion_uses(self):
        rec = record([100.0], [1.0])
        rec['range'] = (6, 25)                                    # channels 7-25
        r = nsz.dls_analyse(rec, nsz.DLS_MODELS[1], 532.0, use_range=True)
        self.assertEqual(len(r['fit']), len(TAU))                 # the fit is still evaluated at every delay
        self.assertAlmostEqual(r['peaks'][0]['median'], 100.0, delta=10.0)
        self.assertEqual(nsz.dls_analyse(rec, nsz.DLS_MODELS[0], 532.0, use_range=False)['p'].shape,
                         nsz.dls_analyse(rec, nsz.DLS_MODELS[0], 532.0, use_range=True)['p'].shape)

    def test_record_carries_the_range_of_the_file(self):
        d = {'tau': TAU * 1e6, 'acf': np.ones(len(TAU)), 'tau_aligned': True,
             'params': {'Angle_deg': 173.0, 'n_solvent': 1.333, 'Visc_mPas': 0.9, 'Temp_C': 25.0,
                        'Range_from': 7, 'Range_to': 33}}
        self.assertEqual(nsz.dls_record(d)['range'], (6, 33))


@unittest.skipUnless(os.path.isfile(os.path.join(LATEX_DIR, "..", "20261118 lipo vuoti varie size DLS",
                                                 "lipo 450 nm 1a40 in PBS_0826.nsz")), "450 nm liposome file not present")
class TestRangeInFile(unittest.TestCase):
    def test_range_of_the_file_matches_the_fit_span_of_the_instrument(self):
        d = nsz.load_nsz(os.path.join(LATEX_DIR, "..", "20261118 lipo vuoti varie size DLS", "lipo 450 nm 1a40 in PBS_0826.nsz"))
        self.assertEqual((d['params']['Range_from'], d['params']['Range_to']), (7, 33))
        self.assertEqual(d['params']['Range_to'], len(d['fit']) + d['fit_offset'])        # Object10 spans the same channels


@unittest.skipUnless(HAS_SCIPY, "scipy not installed")
class TestDlsTab(unittest.TestCase):
    """The tab of the application, on a synthetic sample shaped like load_nsz() output (needs a screen)."""

    @classmethod
    def setUpClass(cls):
        import tkinter as tk
        try:
            cls.root = nsz.TkinterDnD.Tk() if nsz.DND_AVAILABLE else tk.Tk()
        except tk.TclError as e:
            raise unittest.SkipTest("Tk not available: %s" % e)
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls):
        cls.root.destroy()

    def setUp(self):
        self.app = nsz.DLSManager(self.root)
        acf = correlogram([100.0], [1.0])
        d = {'path': 's.nsz', 'name': 's', 'tau': TAU * 1e6, 'acf': acf, 'tau_aligned': True,
             'fit': acf[6:], 'resid': acf[6:] * 0, 'fit_offset': 6, 'sz': np.logspace(0, 3, 84),
             'intens': np.ones(84), 'cumul': np.cumsum(np.ones(84)),
             'params': {'Sample_Name': 's', 'Angle_deg': 173.0, 'n_solvent': 1.333, 'Visc_mPas': 0.8963,
                        'Temp_C': 25.0, 'ZAvg_nm': 100.0, 'PDI': 0.01, 'Mean1_nm': 100.0, 'Mode_nm': 100.0}}
        self.app.samples['s.nsz'] = d
        self.app.listbox.insert('end', 's.nsz')
        self.app.listbox.selection_set(0)
        self.app._on_select()

    def test_recomputed_runs_only_while_the_tab_is_shown_and_compares_with_the_instrument(self):
        app = self.app
        app.size_source.set("Recomputed")
        app._plot_dist()                                               # another tab is shown: nothing is computed
        self.assertEqual(len(app._dls_cache), 0)
        app.notebook.select(app.tab_size)
        self.root.update()
        self.assertGreater(len(app._dls_cache), 0)
        self.assertEqual(len(app.fig_size.axes), 2)
        self.assertIs(app._current_figure(), app.fig_size)
        text = app.dls_text.get("1.0", "end")
        self.assertIn("instrument", text)
        self.assertIn("Z-average", text)

    def test_both_sources_draw_the_same_layout(self):
        app = self.app
        app.notebook.select(app.tab_size)
        for source in ("Instrument", "Recomputed"):
            app.size_source.set(source)
            app._on_source()
            for show in (False, True):
                app.show_cumul.set(show)
                app._plot_dist()
                titles = [ax.get_title() for ax in app.fig_size.axes]
                self.assertEqual(titles[:2], ["Correlogram and fit", "Size distribution by intensity"], source)
                self.assertEqual(len(titles), 3 if show else 2)
                self.assertTrue(app.fig_size.axes[0].lines, source)     # the measured points and the fit are drawn
        app.size_source.set("Instrument")
        app._on_source()
        self.assertIn("Instrument: Z-average", app.dls_text.get("1.0", "end"))

    def test_weighting_changes_the_plot_and_the_text(self):
        self.app.size_source.set("Recomputed"); self.app.notebook.select(self.app.tab_size)
        self.app.dls_weight.set("Number")
        self.app._plot_dls()
        self.assertEqual(self.app.fig_size.axes[1].get_title(), "Size distribution by number")
        self.assertIn("Mie", self.app.dls_text.get("1.0", "end"))             # Mie is the default conversion
        self.app.dls_weight.set("Intensity")
        self.app._plot_dls()
        self.assertIn("instrument mean", self.app.dls_text.get("1.0", "end"))

    def test_mie_and_rayleigh_conversions_in_the_tab(self):
        self.app.size_source.set("Recomputed"); self.app.notebook.select(self.app.tab_size)
        self.app.dls_weight.set("Number")
        self.app.dls_conv.set("Mie")
        self.app.dls_vars['np'].set("1.59")
        self.app._plot_dls()
        self.assertIn("Mie, particle n 1.59", self.app.dls_text.get("1.0", "end"))
        self.app.dls_conv.set("Rayleigh")
        self.app._plot_dls()
        self.assertIn("Rayleigh", self.app.dls_text.get("1.0", "end"))

    def sd_of(self):
        return float(self.app.dls_text.get("1.0", "end").split("S.D. ")[1].split(" nm")[0])

    def test_width_selector_gives_the_instrument_forms_to_every_weighting(self):
        app = self.app
        app.size_source.set("Recomputed"); app.notebook.select(app.tab_size)
        app.dls_model.set(nsz.DLS_MODELS[0])
        app.dls_conv.set("Rayleigh")
        for weight in ("Number", "Intensity"):                         # the intensity distribution takes the form as well
            app.dls_weight.set(weight)
            app.dls_width.set("As computed"); app._plot_dls()
            self.assertNotIn("width set", app.dls_text.get("1.0", "end"))
            computed = self.sd_of()
            sds = {}
            for form, sigma in (("Broad", "0.43"), ("Standard", "0.24"), ("Narrow", "0.07")):
                app.dls_width.set(form); app._plot_dls()
                self.assertIn("width set to σ(ln d) %s as the instrument's %s" % (sigma, form), app.dls_text.get("1.0", "end"))
                sds[form] = self.sd_of()
            self.assertLess(sds["Narrow"], sds["Standard"])
            if weight == "Intensity":      # in number the S.D. in nm is not monotone: a broad peak also moves to smaller sizes
                self.assertLess(sds["Standard"], sds["Broad"])
            self.assertLess(sds["Narrow"], 0.5 * computed)
        app.dls_width.set("As computed")

    def test_manual_resolution_entry_follows_the_form(self):
        app = self.app
        app.size_source.set("Recomputed"); app.notebook.select(app.tab_size)
        app.dls_weight.set("Intensity")
        app.dls_width.set("Standard"); app._on_width()
        self.assertEqual(str(app.ent_resolution.cget("state")), "disabled")
        app.dls_width.set("Manual"); app.dls_vars['resolution'].set("0.6"); app._on_width()
        self.assertEqual(str(app.ent_resolution.cget("state")), "normal")
        self.assertIn("σ(ln d) 0.07 as the instrument's Manual", app.dls_text.get("1.0", "end"))
        app.dls_vars['resolution'].set("0.3"); app._plot_dls()
        self.assertIn("σ(ln d) 0.24", app.dls_text.get("1.0", "end"))
        app.dls_vars['resolution'].set("abc"); app._plot_dls()               # invalid: the Standard resolution
        self.assertIn("σ(ln d) 0.24", app.dls_text.get("1.0", "end"))
        app.dls_vars['resolution'].set("5"); self.assertEqual(app._width_sigma(), nsz.dls_resolution_sigma(1.0))
        app.dls_width.set("As computed"); app._on_width()
        self.assertEqual(str(app.ent_resolution.cget("state")), "disabled")

    def test_both_sources_share_the_diameter_axis(self):
        app = self.app
        app.notebook.select(app.tab_size)
        app.show_cumul.set(True)
        try:
            for scale, lim in (("log", (10.0, 1e4)), ("linear", (0.0, 1000.0))):
                app.size_xscale.set(scale)
                for src in ("Instrument", "Recomputed"):
                    app.size_source.set(src); app._on_source()
                    for ax in app.fig_size.axes[1:]:                   # the distribution and the cumulative panels
                        self.assertEqual(ax.get_xlim(), lim, (scale, src))
        finally:
            app.size_xscale.set("log"); app.show_cumul.set(False)

    def test_guide_is_next_to_the_program_and_opens_from_the_help_menu(self):
        from unittest import mock
        guide = nsz.DLSManager.GUIDE
        self.assertTrue(os.path.isfile(guide))
        with open(guide, encoding="utf-8") as f:
            text = f.read()
        for anchor in ("standard", "setup", "forms", "advanced", "recalc", "dlsmanager", "limits"):
            self.assertIn('id="%s"' % anchor, text)
        self.assertIn("Averaged Mie", text)
        menus = [self.app.root.nametowidget(self.app.root.cget("menu"))]
        labels = [menus[0].entrycget(i, "label") for i in range(menus[0].index("end") + 1)
                  if menus[0].type(i) != "tearoff"]
        self.assertIn("Help", labels)
        with mock.patch("webbrowser.open") as opened:
            self.app.open_guide()
        opened.assert_called_once()
        self.assertTrue(opened.call_args[0][0].startswith("file:///") and opened.call_args[0][0].endswith("sz100_guide.html"))

    def test_standard_form_draws_the_instrument_width(self):
        app = self.app
        app.size_source.set("Recomputed"); app.notebook.select(app.tab_size)
        app.dls_model.set(nsz.DLS_MODELS[0]); app.dls_weight.set("Intensity")
        app.dls_width.set("Standard"); app._plot_dls()
        x, y = app.fig_size.axes[1].lines[0].get_data()
        w = np.asarray(y) / np.sum(y)
        lnd = np.log(np.asarray(x))
        m = np.sum(w * lnd)
        self.assertAlmostEqual(float(np.sqrt(np.sum(w * (lnd - m) ** 2))), 0.24, delta=0.015)
        app.dls_width.set("As computed")

    def test_recomputed_distribution_is_drawn_with_points_like_the_instrument_one(self):
        app = self.app
        app.notebook.select(app.tab_size)
        app.dls_weight.set("Intensity")
        styles = {}
        for src in ("Instrument", "Recomputed"):
            app.size_source.set(src); app._on_source()
            ax = app.fig_size.axes[1]
            ln = max(ax.lines, key=lambda l: len(l.get_xdata()) if l.get_linestyle() == '-' else 0)
            styles[src] = (ln.get_marker(), ln.get_markersize(), ln.get_linewidth())
        self.assertEqual(styles["Instrument"], styles["Recomputed"])
        self.assertEqual(styles["Recomputed"][0], 'o')

    def test_aperture_in_the_tab(self):
        app = self.app
        app.size_source.set("Recomputed"); app.notebook.select(app.tab_size)
        app.dls_weight.set("Volume"); app.dls_conv.set("Mie")
        app.dls_vars['aperture'].set("10")
        app._plot_dls()
        self.assertIn("aperture ±10°", app.dls_text.get("1.0", "end"))
        app.dls_vars['aperture'].set("abc")                            # invalid: the default
        self.assertEqual(app._aperture(), nsz.DLS_APERTURE)
        app.dls_vars['aperture'].set("0")
        self.assertEqual(app._aperture(), 0.0)
        app.dls_vars['aperture'].set("50")                             # limited to 45, not silently replaced by the default
        self.assertEqual(app._aperture(), 45.0)
        app.dls_vars['aperture'].set("-5")
        self.assertEqual(app._aperture(), 0.0)
        app.dls_vars['aperture'].set("10")

    def test_mie_shell_conversion_in_the_tab_with_lamellae(self):
        app = self.app
        app.size_source.set("Recomputed"); app.notebook.select(app.tab_size)
        app.dls_weight.set("Number")
        app.dls_conv.set("Mie shell")
        app.dls_vars['shell_n'].set("7")                               # 7 x 7 nm repeat - 2.5 nm of water = 46.5 nm
        app._plot_dls()
        text = app.dls_text.get("1.0", "end")
        self.assertIn("7 lamellae (shell 46.5 nm", text)
        self.assertAlmostEqual(app._shell()[0], 46.5)
        self.assertAlmostEqual(app._shell()[1], 0.12 * 7 * 4.5 / 46.5)   # the lipid's index averaged over the water gaps
        app.dls_vars['shell_n'].set("abc")                             # invalid: back to one bilayer
        app._plot_dls()
        self.assertIn("1 lamella (shell 4.5 nm", app.dls_text.get("1.0", "end"))
        app.dls_vars['shell_n'].set("1")
        self.assertAlmostEqual(app._shell()[0], 4.5)
        self.assertAlmostEqual(app._shell()[1], 0.12)

    def test_fields_follow_the_conversion(self):
        app = self.app
        entries = {k: w for k, ws in app._conv_fields.items() for w in ws if isinstance(w, tk.Entry)}
        app.dls_conv.set("Mie"); app._on_conv(draw=False)
        states = {k: str(w.cget("state")) for k, w in entries.items()}
        self.assertEqual(states["Mie"], "normal")
        self.assertEqual(states["Mie shell"], "disabled")
        self.assertEqual(states["Mie|Mie shell"], "normal")             # the detector aperture: both Mie conversions
        app.dls_conv.set("Mie shell"); app._on_conv(draw=False)
        self.assertEqual(str(entries["Mie shell"].cget("state")), "normal")
        self.assertEqual(str(entries["Mie"].cget("state")), "disabled")
        self.assertEqual(str(entries["Mie|Mie shell"].cget("state")), "normal")
        app.dls_conv.set("Rayleigh"); app._on_conv(draw=False)
        self.assertEqual({str(w.cget("state")) for w in entries.values()}, {"disabled"})

    def test_rayleigh_warning_for_large_particles(self):
        app = self.app
        app.notebook.select(app.tab_size)
        sample = app.samples['s.nsz']
        sample['sz'] = np.geomspace(1.0, 8000.0, 84)
        big = np.exp(-0.5 * ((np.log(sample['sz']) - np.log(400.0)) / 0.2) ** 2)      # 400 nm
        sample['intens'] = 100.0 * big / big.sum()
        small = np.exp(-0.5 * ((np.log(sample['sz']) - np.log(20.0)) / 0.2) ** 2)     # 20 nm
        app.dls_weight.set("Number")
        app.dls_conv.set("Rayleigh")
        app._plot_size()
        self.assertIn("Rayleigh is more than 10% off above", app.ray_warn.cget("text"))
        app.dls_conv.set("Mie")                                        # Mie: the general caution only
        app._plot_size()
        self.assertNotIn("Rayleigh", app.ray_warn.cget("text"))
        self.assertIn("is not measured", app.ray_warn.cget("text"))
        app.dls_conv.set("Rayleigh")
        app.dls_weight.set("Intensity")                                # intensity needs no caution at all
        app._plot_size()
        self.assertEqual(app.ray_warn.cget("text"), "")
        sample['intens'] = 100.0 * small / small.sum()                 # small particles: Rayleigh is fine
        app.dls_weight.set("Number")
        app._plot_size()
        self.assertNotIn("Rayleigh is more", app.ray_warn.cget("text"))
        self.assertIn("is not measured", app.ray_warn.cget("text"))
        self.assertGreater(app._rayleigh_limit(nsz.dls_record(sample), 532.0, 1.6), 40.0)

    def test_peak_fits_warn_when_they_are_not_credible(self):
        app = self.app
        app.size_source.set("Recomputed"); app.notebook.select(app.tab_size)
        app.dls_model.set(nsz.DLS_MODELS[2])                           # 2 peaks on a single-size sample
        app._plot_dls()
        text = app.ray_warn.cget("text")
        self.assertIn("second peak is not supported", text)
        app.dls_model.set(nsz.DLS_MODELS[0])                           # free model, intensity: nothing to warn about
        app.dls_weight.set("Intensity")
        app._plot_dls()
        self.assertEqual(app.ray_warn.cget("text"), "")

    def test_the_recomputed_fit_line_spans_only_the_points_that_were_used(self):
        app = self.app
        app.size_source.set("Recomputed"); app.notebook.select(app.tab_size)
        app._plot_dist()
        fit = [ln for ln in app.fig_size.axes[0].lines if getattr(ln, "_cursor_name", "").endswith("fit")][0]
        sample = app.samples['s.nsz']
        tau, g1 = sample['tau'], np.sqrt(np.clip(sample['acf'], 0, None))
        x = np.asarray(fit.get_xdata())
        self.assertGreaterEqual(x.min(), 1.0)                          # delays below 1 us are not used
        self.assertLess(x.max(), tau.max())                            # nor the tail below g1 min: not the whole axis
        self.assertTrue(np.all(g1[np.isin(tau, x)] > 0.1 - 1e-9))

    def test_the_source_selector_swaps_the_two_plots(self):
        app = self.app
        app.notebook.select(app.tab_size)
        self.assertEqual(app.box_rec.winfo_manager(), "")              # the model controls belong to Recomputed
        self.assertIs(app._current_figure(), app.fig_size)
        app.size_source.set("Recomputed")
        app._on_source()
        self.root.update()
        self.assertEqual(app.box_rec.winfo_manager(), "pack")
        self.assertIs(app._current_figure(), app.fig_size)             # the same figure for both sources
        app.size_source.set("Instrument")
        app._on_source()
        self.assertEqual(app.box_rec.winfo_manager(), "")
        self.assertEqual(len(app.notebook.tabs()), 3)                  # Autocorrelogram, Size distribution, Parameters

    def test_instrument_distribution_by_volume_and_number(self):
        app = self.app
        app.notebook.select(app.tab_size)
        for weight in ("Volume", "Number"):
            for conv in ("Mie", "Mie shell", "Rayleigh"):
                app.dls_weight.set(weight)
                app.dls_conv.set(conv)
                app._plot_dist()
                ax = app.fig_size.axes[1]
                self.assertEqual(ax.get_ylabel(), weight + "-weighted (normalised)")
                dist = [ln for ln in ax.lines if len(ln.get_xdata()) > 10][0]
                self.assertAlmostEqual(float(np.max(dist.get_ydata())), 1.0)

    def test_models_and_invalid_settings(self):
        self.app.size_source.set("Recomputed"); self.app.notebook.select(self.app.tab_size)
        self.app.dls_model.set(nsz.DLS_MODELS[2])                      # also sets g1 min to 0.05
        self.assertEqual(self.app.dls_vars['gmin'].get(), '0.05')
        self.assertIn("Peak 2", self.app.dls_text.get("1.0", "end"))
        self.assertIn("BIC", self.app.dls_text.get("1.0", "end"))
        self.app.dls_vars['wl'].set('abc')
        self.app._plot_dls()
        self.assertIn("Invalid settings", self.app.dls_text.get("1.0", "end"))


if __name__ == "__main__":
    unittest.main()
