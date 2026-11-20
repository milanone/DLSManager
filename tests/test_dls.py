"""Tests for the DLS analysis of the stored correlogram: synthetic correlograms of known sizes (no files, no Tk),
and the real sample in 'example data/' against the values of the instrument (skipped if the file is missing).

    py -m unittest discover -s tests -v
"""
import importlib.util
import os
import unittest

import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("nsz", os.path.join(HERE, "..", "nsz_viewer.pyw"))
nsz = importlib.util.module_from_spec(spec)
spec.loader.exec_module(nsz)

SAMPLE = os.path.join(HERE, "..", "example data", "lipo_bup_MLV_0minson_scattering_0729.nsz")

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


@unittest.skipUnless(os.path.isfile(SAMPLE) and HAS_SCIPY, "real sample or scipy not present")
class TestRealSample(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.d = nsz.load_nsz(SAMPLE)
        cls.rec = nsz.dls_record(cls.d)

    def test_record_uses_the_conditions_of_the_file(self):
        self.assertEqual(self.rec['angle'], 173.0)
        self.assertEqual(len(self.rec['tau']), 43)

    def test_z_average_is_close_to_the_instrument(self):
        # This sample is very polydisperse (PI 0.6): the cumulant fit is least reliable here; 1-2% on less
        # polydisperse files of LabSpectrumManager, 4% on this one.
        r = nsz.dls_analyse(self.rec, nsz.DLS_MODELS[0], self.d['params']['Lambda_nm'])
        self.assertAlmostEqual(r['z'] / self.d['params']['ZAvg_nm'], 1.0, delta=0.06)

    def test_all_models_run(self):
        for m in nsz.DLS_MODELS:
            r = nsz.dls_analyse(self.rec, m, 532.0)
            self.assertTrue(np.all(np.isfinite(r['p'])))
            self.assertEqual(len(r['fit']), 43)


LATEX_DIR = os.path.join(HERE, "..", "..", "LabSpectrumManager", "spettri di esempio", "20261119 DLS lattice e liposomi")
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
        self.app = nsz.NSZViewer(self.root)
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

    def test_runs_only_on_its_tab_and_compares_with_the_instrument(self):
        self.assertEqual(len(self.app.fig_dls.axes), 0)               # not computed while another tab is shown
        self.app.notebook.select(self.app.tab_dls)
        self.root.update()
        self.assertEqual(len(self.app.fig_dls.axes), 2)
        self.assertIs(self.app._current_figure(), self.app.fig_dls)
        text = self.app.dls_text.get("1.0", "end")
        self.assertIn("instrument", text)
        self.assertIn("Z-average", text)

    def test_models_and_invalid_settings(self):
        self.app.notebook.select(self.app.tab_dls)
        self.app.dls_model.set(nsz.DLS_MODELS[2])                      # also sets g1 min to 0.05
        self.assertEqual(self.app.dls_vars['gmin'].get(), '0.05')
        self.assertIn("Peak 2", self.app.dls_text.get("1.0", "end"))
        self.assertIn("BIC", self.app.dls_text.get("1.0", "end"))
        self.app.dls_vars['wl'].set('abc')
        self.app._plot_dls()
        self.assertIn("Invalid settings", self.app.dls_text.get("1.0", "end"))


if __name__ == "__main__":
    unittest.main()
