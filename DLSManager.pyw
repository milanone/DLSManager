"""
DLS Manager — Horiba SZ-100 DLS: .nsz viewer, correlogram analysis and size distributions
Supports drag & drop, multiple-sample comparison, Excel export.
The plots follow the shared Origin style of PlotStyleKit when the sibling repo is present.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, struct, re, csv, io, pickle, traceback
try:
    import olefile
    OLEFILE_AVAILABLE = True
except ImportError:
    OLEFILE_AVAILABLE = False
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    XLSX_AVAILABLE = True
except ImportError:
    XLSX_AVAILABLE = False

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import numpy as np

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False


# ---- PlotStyleKit bootstrap: the same text in every program (PlotStyleKit/README.md, 'Using it from a program') ----
import json
import os

PSK_MIN_API = 2
PSK_CONFIG = os.path.join(os.path.expanduser('~'), '.plotstylekit.json')


def _psk_config():
    try:
        with open(PSK_CONFIG, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _psk_save(**changes):
    cfg = _psk_config()
    cfg.update(changes)
    try:
        with open(PSK_CONFIG, 'w', encoding='utf-8') as f:
            json.dump(cfg, f)
    except OSError:
        pass


def _psk_try(folder, min_api):
    """(module, problem): problem is None when plotstylekit.py in `folder` loads and is recent enough."""
    import importlib.util as ilu
    path = os.path.join(folder, 'plotstylekit.py')
    if not os.path.isfile(path):
        return None, f"plotstylekit.py is not in {folder}"
    try:
        spec = ilu.spec_from_file_location('plotstylekit', path)
        mod = ilu.module_from_spec(spec)
        spec.loader.exec_module(mod)
    except Exception as e:
        return None, f"plotstylekit.py in {folder} cannot be loaded: {e}"
    if getattr(mod, 'API', 0) < min_api:
        return None, (f"PlotStyleKit in {folder} is too old for this program (API {getattr(mod, 'API', '?')}, "
                      f"it needs {min_api}): update it")
    return mod, None


def _load_plotstylekit(min_api=PSK_MIN_API):
    """(module, problem): the first usable PlotStyleKit, or None and a sentence saying what was wrong."""
    here = os.path.dirname(os.path.abspath(__file__))
    folders = [here, os.environ.get('PLOTSTYLEKIT'), _psk_config().get('path'), os.path.join(here, '..', 'PlotStyleKit'),
               os.path.expanduser(os.path.join('~', 'progetti', 'PlotStyleKit'))]
    problem = "PlotStyleKit was not found (next to the program, PLOTSTYLEKIT, a folder you chose, ../PlotStyleKit)"
    for folder in filter(None, folders):
        if os.path.isfile(os.path.join(folder, 'plotstylekit.py')):
            mod, why = _psk_try(folder, min_api)
            if mod is not None:
                return mod, None
            problem = why
    return None, problem


def _ask_plotstylekit_folder(parent, problem, min_api=PSK_MIN_API):
    """The warning that asks for the folder of PlotStyleKit. Returns the module, or None if the user cancels
    (remembered: the program is not asked again) or closes the folder dialog (asked again next time)."""
    from tkinter import filedialog, messagebox
    while True:
        if not messagebox.askokcancel(
                "PlotStyleKit", problem + ".\n\nOK: choose the folder that contains plotstylekit.py.\n"
                "Cancel: use the program without it (plain matplotlib look, no Edit Figure). You will not be asked "
                "again; File > PlotStyleKit... lets you change your mind.", parent=parent):
            _psk_save(declined=True)
            return None
        folder = filedialog.askdirectory(parent=parent, title="Folder of PlotStyleKit (contains plotstylekit.py)")
        if not folder:
            return None
        mod, why = _psk_try(folder, min_api)
        if mod is not None:
            _psk_save(path=folder, declined=False)
            return mod
        problem = why


psk, PSK_PROBLEM = _load_plotstylekit()
if psk is not None:
    try:
        psk.setup()
    except Exception as e:
        psk, PSK_PROBLEM = None, f"PlotStyleKit failed to start: {e}"
# ---- end of the PlotStyleKit bootstrap ----

# ── Colour palette for multiple samples (used only without PlotStyleKit) ─────
COLORS = ["#2563EB","#DC2626","#16A34A","#D97706","#7C3AED",
          "#0891B2","#BE185D","#65A30D","#EA580C","#4338CA","#0D9488"]
MAX_SAMPLES_SHOWN = len(COLORS)


def sample_colors():
    """Colours for the overlaid samples: the Origin cycle (black first) if PlotStyleKit is loaded."""
    if psk is not None:
        cycle = matplotlib.rcParams['axes.prop_cycle'].by_key().get('color')
        if cycle:
            return list(cycle)
    return COLORS

# ── NSZ parser ────────────────────────────────────────────────────────────────

def f32(data):
    n = len(data)//4
    return list(struct.unpack(f'<{n}f', data[:n*4]))

def parse_props(data):
    props = {}
    i = 0
    while i < len(data) - 10:
        j = data.find(b'\x00', i)
        if j == -1 or j - i < 2:
            i += 1; continue
        key_b = data[i:j]
        if all(32 <= b < 127 for b in key_b) and len(key_b) >= 3:
            if j + 5 < len(data) and data[j+5] == 0x00:
                try:
                    v = struct.unpack('<f', data[j+1:j+5])[0]
                    if not (v != v):
                        props[key_b.decode('ascii')] = v
                except (struct.error, UnicodeDecodeError): pass
            i = j + 1
        else:
            i += 1
    return props

def _open_ole(path):
    if not OLEFILE_AVAILABLE:
        raise ImportError("The olefile library is not available.\nInstall it with: pip install olefile")
    return olefile.OleFileIO(path)

def dump_nsz(path):
    """Print the internal structure of an .nsz file for debugging (time base, correlator parameters)."""
    ole = _open_ole(path)
    try:
        def s(name):
            return ole.openstream(name).read() if ole.exists(name) else b''

        print(f"\n=== {os.path.basename(path)} ===")
        print("Object3 header (first 16 float32):")
        hdr = f32(s('Object3'))[:16]
        for i, v in enumerate(hdr):
            print(f"  [{i:2d}] {v:.6g}")

        print("\nBase34 all float parameters:")
        p34 = parse_props(s('Base34'))
        for k, v in sorted(p34.items()):
            print(f"  {k} = {v:.6g}")
    finally:
        ole.close()

def load_nsz(path):
    """Load an .nsz file and return a dict with all the data."""
    ole = _open_ole(path)
    try:
        return _read_nsz(path, ole)
    finally:
        ole.close()

def _nsz_correlogram(ole):
    """(delay times in µs, normalised g2-1) from the float32 arrays 'Delay' (seconds) and 'OriginalACF', or None.
    Each array sits in one stream, preceded by its byte length (u32) and its key name; the ACF has a few junk values
    after the real ones, cut to the number of delays. Checked on the sample (43 channels, 0.1 µs to 0.109 s) and
    on four other SZ-100 files in LabSpectrumManager."""
    def arr(b, key):
        m = re.search(re.escape(key) + b'\x00', b)
        if not m or m.start() < 4:
            return None
        n = struct.unpack('<I', b[m.start() - 4:m.start()])[0]
        if n % 4 or not 16 <= n <= 4 * 4096 or m.end() + n > len(b):
            return None
        return np.array(struct.unpack('<%df' % (n // 4), b[m.end():m.end() + n]), float)
    for entry in ole.listdir(streams=True, storages=False):
        b = ole.openstream(entry).read()
        tau, acf = arr(b, b'Delay'), arr(b, b'OriginalACF')
        if tau is None or acf is None or len(acf) < len(tau):
            continue
        acf = acf[:len(tau)]
        if (np.all(np.isfinite(tau)) and np.all(tau > 0) and np.all(np.diff(tau) > 0) and np.all(np.isfinite(acf))
                and acf.max() <= 1.5 and acf.min() >= -0.5):
            return tau * 1e6, acf
    return None


def _nsz_int_key(ole, key):
    """int32 stored after `key` and a NUL byte in any stream, or None."""
    for entry in ole.listdir(streams=True, storages=False):
        m = re.search(re.escape(key) + b'\x00(.{4})', ole.openstream(entry).read(), re.S)
        if m:
            return struct.unpack('<i', m.group(1))[0]
    return None


def _nsz_curves(ole, acf, ch_from, ch_to):
    """Size distribution and fit curves found by CONTENT, because the software versions number the `Object` streams
    differently (2026-06 file: sizes Object5, intensity Object4, cumulative Object31, fit Object10, residuals Object7;
    2026-11-19 files: Object6, Object5, Object32, Object11, Object8). Returns a dict with whichever of 'sz' (17 header
    floats, 84 diameters ending near 8510 nm), 'intens' (3 header floats, 84 values summing to 100), 'cumul' (4 header
    floats, 84 values rising to 100), 'fit' and 'resid' (3 header floats, then one value per correlator channel
    ch_from..ch_to, g2-1) it could identify. The fit is the stream of that length closest to the measured ACF over the
    channels, the residual the small-valued one most correlated with ACF - fit."""
    arrs = {}
    for entry in ole.listdir(streams=True, storages=False):
        if len(entry) == 1 and entry[0].startswith('Object'):
            raw = ole.openstream(entry).read()
            if len(raw) > 2000:          # a 100 kB blob (not a curve) in the newer files
                continue
            arrs[entry[0]] = np.frombuffer(raw[:len(raw) // 4 * 4], '<f4').astype(float)
    names = sorted(arrs, key=lambda n: int(n[6:]) if n[6:].isdigit() else 0)
    out = {}
    for n in names:
        a = arrs[n]
        if not np.all(np.isfinite(a)):
            continue
        if len(a) == 101 and 8000 < a[-1] < 9000:
            out.setdefault('sz', a[17:101])
        elif len(a) == 88:
            if abs(a[87] - 100) < 0.5 and np.all(np.diff(a[4:88]) >= -1e-3):
                out.setdefault('cumul', a[4:88])
            elif abs(a[3:87].sum() - 100) < 1.0:
                out.setdefault('intens', a[3:87])
    if ch_from and ch_to:
        n_ch = ch_to - ch_from + 1
        ref = np.asarray(acf[ch_from - 1:ch_to], float)
        cands = [arrs[n][3:] for n in names if len(arrs[n]) == 3 + n_ch and np.all(np.isfinite(arrs[n]))]
        if len(ref) == n_ch and cands:
            fits = [c for c in cands if 0.3 < c.max() < 1.2]
            if fits:
                fit = min(fits, key=lambda c: float(np.sqrt(np.mean((c - ref) ** 2))))
                out['fit'] = fit
                small = [c for c in cands if np.abs(c).max() < 0.05 and c.std() > 0]
                if small:
                    out['resid'] = max(small, key=lambda c: abs(float(np.corrcoef(c, ref - fit)[0, 1])))
    return out


def _plausible_fit(fit, acf, offset):
    """True if `fit` (g2-1 at the channels from `offset`) follows the measured `acf` there: at least 5 values, between 0
    and 1.2, and within 0.1 rms of it. Stops a stream of another layout (e.g. a 16-float header) from being drawn as a fit."""
    fit = np.asarray(fit, float)
    ref = np.asarray(acf[offset:offset + len(fit)], float)
    return (len(fit) >= 5 and len(ref) == len(fit) and np.all(np.isfinite(fit)) and fit.max() <= 1.2
            and fit.min() >= -0.1 and float(np.sqrt(np.mean((fit - ref) ** 2))) < 0.1)


def _read_nsz(path, ole):
    def s(name):
        return ole.openstream(name).read() if ole.exists(name) else b''

    # ACF: all 43 channels, with their delays, from the 'Delay' / 'OriginalACF' arrays (see _nsz_correlogram).
    # Object3 lacks the first 6 delays (0.1-0.7 us) while Object2 starts at channel 0, so pairing them
    # shifts the x axis by 6 channels and drops the last 6 real points: it is only a fallback.
    corr = _nsz_correlogram(ole)
    tau_aligned = corr is not None
    if tau_aligned:
        tau, acf = corr
    else:
        tau_full = f32(s('Object3'))[16:]
        acf_full = f32(s('Object2'))[3:]
        cut = next((i for i in range(1, len(tau_full)) if tau_full[i] <= tau_full[i-1]),
                   len(tau_full))
        tau = tau_full[:cut]
        acf = acf_full[:cut]

    ch_from, ch_to = _nsz_int_key(ole, b'Dist_ChannelFrom'), _nsz_int_key(ole, b'Dist_ChannelTo')
    if not (ch_from and ch_to and 1 <= ch_from < ch_to <= len(tau)):
        ch_from = ch_to = None
    curves = _nsz_curves(ole, acf, ch_from, ch_to)

    # Horiba fit and residuals.
    # Object10 header: [int=1][int=n_fit][int=fit_offset], then the fit values.
    # fit_offset = number of initial channels (afterpulsing) excluded from the analysis.
    obj10_raw = f32(s('Object10'))
    FIT_OFFSET = struct.unpack('<I', struct.pack('<f', obj10_raw[2]))[0] if len(obj10_raw) > 2 else 6
    fit   = obj10_raw[3:]
    resid = f32(s('Object7'))[3:]
    if 'fit' in curves:      # found by content: the Object numbers depend on the software version
        FIT_OFFSET, fit = ch_from - 1, list(curves['fit'])
        resid = list(curves['resid']) if 'resid' in curves else []
    elif not _plausible_fit(fit, acf, FIT_OFFSET):
        fit, resid = [], []      # Object10 of another software version: better no fit than a wrong one

    # Size distribution: 84 bins, with a different header length in each stream (17, 3 and 4 float32);
    # validated against the CSV exported by the instrument software (all three match to its rounding)
    sz    = f32(s('Object5'))[17:17+84]
    intens= f32(s('Object4'))[3:3+84]
    cumul = f32(s('Object31'))[4:4+84]
    if all(k in curves for k in ('sz', 'intens', 'cumul')):
        sz, intens, cumul = list(curves['sz']), list(curves['intens']), list(curves['cumul'])
    nb    = min(len(sz), len(intens), len(cumul))
    sz, intens, cumul = sz[:nb], intens[:nb], cumul[:nb]

    # Parameters
    p1  = parse_props(s('Base1'))
    # Measurement parameters and sample name sit in Base34 / Base38 in some software versions and in Base35 / Base39
    # in others (files of 2026-11-19): pick the stream by its content, not by its number.
    def by_key(key, default):
        for entry in ole.listdir(streams=True, storages=False):
            if len(entry) == 1 and entry[0].startswith('Base'):
                data = ole.openstream(entry).read()
                if key in data:
                    return data
        return s(default)
    d34 = by_key(b'MeasAngle', 'Base34')
    d38 = by_key(b'Sample_Name', 'Base38')
    p34 = parse_props(d34)

    def get_str(data, key):
        m = re.search(key + b'\x00([\\x20-\\x7e]+)', data)
        return m.group(1).decode().strip() if m else ''

    m = re.search(rb'MeasID\x00{0,4}(\d{15})', d34)
    raw_id = m.group(1).decode() if m else '000000000000000'

    def pv(d, k, scale=1):
        v = d.get(k)
        if v is not None and abs(v) < 1e6:
            return round(v * scale, 4)
        return None

    params = {
        'Sample_Name': get_str(d38, b'Sample_Name') or os.path.splitext(os.path.basename(path))[0],
        'Date': f'{raw_id[0:4]}-{raw_id[4:6]}-{raw_id[6:8]}',
        'Time': f'{raw_id[8:10]}:{raw_id[10:12]}:{raw_id[12:14]}',
        'Solvent': get_str(d34, b'MeasSolventName'),
        'Software': get_str(d34, b'AppVersion'),
        'Angle_deg': pv(p34, 'MeasAngle'),
        'Lambda_nm': pv(p34, 'MeasWaveLength'),
        'Temp_C':   pv(p34, 'MeasHolderTemp'),
        'Visc_mPas':pv(p34, 'MeasSolventVisco'),
        'n_solvent': pv(p34, 'MeasSolventRef'),
        # Correlator channels (1-based, inclusive) the instrument used for the size distribution; None if absent
        'Range_from': ch_from,
        'Range_to': ch_to,
        'n_particle': pv(p34, 'MeasSampleRef'),   # index entered in the instrument software (1.6 default)
        # CalcMean/SD/Mode/PeakPos/D10/D50/D90 are stored as RADII → ×2
        # PDI: CalcTotalPI is wrong, use Cum_fPi
        'Peak1_nm': pv(p1, 'CalcPeakPos[0]', 2.0),
        'Mean1_nm': pv(p1, 'CalcMean[0]',    2.0),
        'SD1_nm':   pv(p1, 'CalcSD[0]',      2.0),
        'Area1_pct':pv(p1, 'CalcArea[0]'),
        'Peak2_nm': pv(p1, 'CalcPeakPos[1]', 2.0),
        'Mean2_nm': pv(p1, 'CalcMean[1]',    2.0),
        'SD2_nm':   pv(p1, 'CalcSD[1]',      2.0),
        'D10_nm':   pv(p1, 'CalcPerDiameter[0]', 2.0),
        'D50_nm':   pv(p1, 'CalcPerDiameter[4]', 2.0),
        'D90_nm':   pv(p1, 'CalcPerDiameter[8]', 2.0),
        'Mode_nm':  pv(p1, 'CalcMode',   2.0),
        'Median_nm':pv(p1, 'CalcMedian', 2.0),
        'ZAvg_nm':  pv(p1, 'Cum_fMean',  2.0),
        'PDI':      pv(p1, 'Cum_fPi'),          # Cum_fPi = the correct PDI
        'Span':     pv(p1, 'CalcSpan'),
        'GeoMean_nm': pv(p1, 'CalcGeoMean', 2.0),
        'AriMean_nm': pv(p1, 'CalcAriMean', 2.0),
    }

    return {
        'path': path,
        'name': os.path.splitext(os.path.basename(path))[0],
        'tau':        np.array(tau,   dtype=float),
        'acf':        np.array(acf,   dtype=float),
        'fit':        np.array(fit,   dtype=float),
        'resid':      np.array(resid, dtype=float),
        'fit_offset': FIT_OFFSET,
        'tau_aligned': tau_aligned,   # False: legacy Object3/Object2 pairing, delays not matched to the channels
        'sz':        np.array(sz,    dtype=float),
        'intens':     np.array(intens,dtype=float),
        'cumul':      np.array(cumul, dtype=float),
        'params':     params,
    }

class ToolTip:
    """Small yellow help box shown after the pointer rests on a widget for `delay` ms; hidden on leave or click."""

    def __init__(self, widget, text, delay=500, wraplength=340):
        self.widget, self.text, self.delay, self.wraplength = widget, text, delay, wraplength
        self._after, self._tip = None, None
        widget.bind('<Enter>', self._schedule, add='+')
        widget.bind('<Leave>', self._hide, add='+')
        widget.bind('<ButtonPress>', self._hide, add='+')

    def _schedule(self, _event=None):
        self._hide()
        self._after = self.widget.after(self.delay, self._show)

    def _show(self):
        self._after = None
        if self._tip is not None or not self.text:
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self._tip, text=self.text, justify=tk.LEFT, bg='#ffffe0', relief=tk.SOLID, borderwidth=1,
                 font=('Arial', 8), wraplength=self.wraplength).pack()

    def _hide(self, _event=None):
        if self._after is not None:
            self.widget.after_cancel(self._after)
            self._after = None
        if self._tip is not None:
            self._tip.destroy()
            self._tip = None


# ── DLS analysis of the stored correlogram ────────────────────────────────────
# Ported from LabSpectrumManager, where it was checked on four SZ-100 files (532 nm, 43 delays from 1e-7 to 0.109 s):
# the Z-average of the cumulants is within 1-2% of the instrument's; the two-peak split is an ill-posed inversion,
# not stable between replicates, and only meaningful for sizes that differ by about 3 times or more.

_KB = 1.380649e-23
DLS_MODELS = ("Regularised (free distribution)", "1 peak (log-normal)", "2 peaks (log-normal)")
DLS_TIPS = {
    'model': "Regularised: a free size distribution (non-negative, smoothed), the least assumptions but a limited "
             "resolution. 1 or 2 peaks: the intensity distribution is one or two log-normal peaks fitted to the "
             "correlogram (diameter of each peak, its width and its share of the scattered intensity). Two peaks are "
             "only meaningful if the sizes differ by a factor of about 3 or more and the fit is clearly better than "
             "with one.",
    'wl': "Laser wavelength, nm (532 for the SZ-100; the sizes depend on it through the scattering vector).",
    'alpha': "Smoothing strength of the regularised inversion (second-difference penalty): larger = smoother and wider "
             "peaks, smaller = more detail and more spurious peaks. Only used by the regularised model.",
    'weight': "Intensity: as fitted. Volume and number: the same distribution divided by d^3 and d^6 (Rayleigh "
              "approximation, I ~ N d^6, as in the Distribution tab). Exact for particles much smaller than the "
              "wavelength in the medium (400 nm); above about 100 nm the instrument uses Mie theory with the particle "
              "refractive index, so its volume and number values differ. The shift is large for broad or "
              "multi-peak distributions: a small tail in intensity can be most of the particles.",
    'conv': "How volume and number are obtained from the intensity distribution. Mie: the light one particle scatters "
            "at the measurement angle, from Mie theory with the refractive index of the particle (the one stored "
            "in the file by the instrument software, 1.6 by default, or the value typed here; leave empty for the "
            "file's). Mie shell: a vesicle (liposome) as a shell around a core of medium, with the shell thickness and "
            "index excess typed below. Rayleigh: scattered intensity ~ d^6, valid only well below 100 nm.",
    'tail': "Free (regularised) distribution only, for volume and number: bins below this percentage of the maximum "
            "are dropped before the conversion. The inversion leaves a few percent of intensity at the edge of the "
            "size grid, which the conversion would turn into most of the particles. 0 keeps everything.",
    'range': "Use only the correlator channels that the instrument software used for the size distribution (stored in "
             "the file: from channel 7 up to a channel that depends on the measurement). Without it, every point "
             "above the g1 minimum is used. Files without that information use all points. The cumulants are not "
             "restricted.",
    'shell': "Mie shell: a vesicle as a shell around a core filled with the medium. Thickness of the shell in nm (a "
             "single bilayer is about 4.5 nm; each extra lamella adds about as much) and its refractive index above "
             "the medium's (0.12 for a lipid bilayer, an order-of-magnitude value). With the outer diameter from the "
             "distribution, the light one vesicle scatters grows about as d^4 at small sizes instead of d^6.",
    'gmin': "Only the points of the correlogram with g1 = sqrt(ACF) above this value are used: the tail is mostly "
            "noise. 0.1 for the regularised model, 0.05 for the peaks.",
}


def dls_q2(angle_deg, n_solvent, wavelength_nm=532.0):
    """Squared scattering vector in 1/m^2: q = 4 pi n sin(theta/2) / lambda."""
    q = 4 * np.pi * n_solvent * np.sin(np.radians(angle_deg) / 2) / (wavelength_nm * 1e-9)
    return q * q


def _dls_gamma(d_nm, q2, temp_c, visc_mpas):
    """Decay rate (1/s) of the field correlation of spheres of diameter d_nm: Gamma = D q^2, D = kT / (3 pi eta d)."""
    return _KB * (temp_c + 273.15) / (3 * np.pi * visc_mpas * 1e-3 * np.asarray(d_nm, float) * 1e-9) * q2


def dls_cumulants(tau, acf, q2, temp_c, visc_mpas, gmin=0.30):
    """Z-average (nm) and polydispersity index from a quadratic fit of ln g1 (g1 = sqrt of the normalised ACF) over the
    points with g1 > gmin: ln g1 = -G t + (mu2/2) t^2, Z from G, PI = mu2 / G^2. `tau` in seconds. Returns None when
    there are too few points."""
    tau, acf = np.asarray(tau, float), np.asarray(acf, float)
    g1 = np.sqrt(np.clip(acf, 0, None))
    sel = g1 > gmin
    if sel.sum() < 4:
        return None
    ts = 1e-4
    c = np.polyfit(tau[sel] / ts, np.log(g1[sel]), 2)
    G, mu2 = -c[1] / ts, 2 * c[0] / ts ** 2
    if G <= 0:
        return None
    return float(_KB * (temp_c + 273.15) / (3 * np.pi * visc_mpas * 1e-3 * (G / q2)) * 1e9), float(mu2 / G ** 2)


def dls_invert(tau, acf, q2, temp_c, visc_mpas, alpha=1.0, gmin=0.1, dmin=10.0, dmax=3000.0, ngrid=100, tau_min=1e-6,
               tau_max=np.inf):
    """Intensity-weighted size distribution by regularised non-negative least squares: g1 = sum_j p_j exp(-Gamma_j t)
    with p >= 0 on a log grid of diameters, weighted by g1 and a second-difference smoothing penalty alpha. Returns
    (diameters nm, p normalised to sum 1, fitted g1 at all delays). Needs scipy (nnls)."""
    from scipy.optimize import nnls
    tau, acf = np.asarray(tau, float), np.asarray(acf, float)
    g1 = np.sqrt(np.clip(acf, 0, None))
    d = np.geomspace(dmin, dmax, ngrid)
    gam = _dls_gamma(d, q2, temp_c, visc_mpas)
    sel = (g1 > gmin) & (tau >= tau_min) & (tau <= tau_max)
    K = np.exp(-np.outer(tau[sel], gam))
    w = g1[sel]
    L = np.diff(np.eye(ngrid), 2, axis=0)
    p, _ = nnls(np.vstack([K * w[:, None], alpha * L]), np.concatenate([g1[sel] * w, np.zeros(ngrid - 2)]))
    p = p / p.sum()
    return d, p, np.exp(-np.outer(tau, gam)) @ p


def dls_fit_peaks(tau, acf, q2, temp_c, visc_mpas, n_peaks=2, gmin=0.05, dmin=10.0, dmax=1e4, ngrid=160,
                  tau_min=1e-6, tau_max=np.inf):
    """Intensity distribution as n_peaks log-normal peaks (median diameter, width in ln d, share of the intensity),
    fitted to g1 by least squares from several starting points (the intercept of g1 is a free parameter). Returns
    (diameters, p, fitted g1, peaks, rms of g1, BIC) where peaks is a list of dicts {'median', 'sigma', 'fraction'} sorted
    by size and BIC = N ln(rss/N) + k ln N is the Bayesian information criterion (a model with the lower BIC is
    preferred; a difference below about 6 is not significant). Needs scipy (least_squares).
    Delays shorter than tau_min (1 us) are left out: the first channels of the stored correlogram drop by about 1% in
    one step (the first point is the normalisation), which a fit reads as a 3-5% 'peak' at 3-16 nm on every file."""
    from scipy.optimize import least_squares
    tau, acf = np.asarray(tau, float), np.asarray(acf, float)
    g1 = np.sqrt(np.clip(acf, 0, None))
    d = np.geomspace(dmin, dmax, ngrid)
    lnd = np.log(d)
    gam = _dls_gamma(d, q2, temp_c, visc_mpas)
    sel = (g1 > gmin) & (tau >= tau_min) & (tau <= tau_max)
    K = np.exp(-np.outer(tau[sel], gam))
    w, y = g1[sel], g1[sel]
    n = n_peaks

    def unpack(x):
        mus, sig, logit = x[:n], x[n:2 * n], x[2 * n:3 * n]
        a = np.exp(logit - logit.max())
        a /= a.sum()
        p = np.zeros(ngrid)
        for k in range(n):
            g = np.exp(-0.5 * ((lnd - mus[k]) / sig[k]) ** 2)
            p += a[k] * g / g.sum()
        return p, a, mus, sig

    def resid(x):
        return w * (x[3 * n] * (K @ unpack(x)[0]) - y)         # x[3n] = intercept of g1 (the stored ACF starts near 0.99)

    cum = dls_cumulants(tau, acf, q2, temp_c, visc_mpas)
    z0 = cum[0] if cum else 100.0
    starts = [[np.log(z0)]] if n == 1 else [[np.log(z0 / f1), np.log(z0 * f2)] for f1, f2 in ((1.3, 3), (1.6, 6), (2, 4))]
    lo = np.r_[[np.log(dmin) + 0.5] * n, [0.05] * n, [-8.0] * n, 0.5]
    hi = np.r_[[np.log(dmax) - 0.5] * n, [1.2] * n, [8.0] * n, 1.5]
    best = None
    for mu0 in starts:
        x0 = np.r_[mu0, [0.15] * n, [0.0] * n, 1.0]
        res = least_squares(resid, np.clip(x0, lo + 1e-9, hi - 1e-9), bounds=(lo, hi))
        if best is None or res.cost < best.cost:
            best = res
    p, a, mus, sig = unpack(best.x)
    order = np.argsort(mus)
    peaks = [{'median': float(np.exp(mus[k])), 'sigma': float(sig[k]), 'fraction': float(a[k])} for k in order]
    rss, npts, k = float(np.sum(resid(best.x) ** 2)), int(sel.sum()), 3 * n
    bic = npts * np.log(max(rss, 1e-30) / npts) + k * np.log(npts)         # lower is better; compare 1 with 2 peaks
    return (d, p, best.x[3 * n] * (np.exp(-np.outer(tau, gam)) @ p), peaks, float(np.sqrt(rss / npts)), float(bic))


def dls_analyse(rec, model=DLS_MODELS[0], wavelength=532.0, alpha=1.0, gmin=None, use_range=False):
    """Z-average, PI and the intensity distribution of a stored correlogram `rec` ({'tau' in seconds, 'acf', 'angle',
    'n', 'visc', 'T'}) with the chosen model of DLS_MODELS. Returns a dict with z, pi (cumulants, may be None), d, p,
    fit, peaks (None for the regularised model), mean, mode, sd of the distribution."""
    q2 = dls_q2(rec['angle'], rec['n'], wavelength)
    args = (rec['tau'], rec['acf'], q2, rec['T'], rec['visc'])
    cum = dls_cumulants(*args)
    out = {'z': cum[0] if cum else None, 'pi': cum[1] if cum else None, 'peaks': None}
    # delays the inversion may use: those of the channels the instrument used (cumulants are not restricted)
    lim = {}
    if use_range and rec.get('range'):
        lo, hi = rec['range']
        lim = {'tau_min': rec['tau'][lo] * (1 - 1e-9), 'tau_max': rec['tau'][hi - 1] * (1 + 1e-9)}
    if model == DLS_MODELS[0]:
        out['d'], out['p'], out['fit'] = dls_invert(*args, alpha=alpha, gmin=0.1 if gmin is None else gmin, **lim)
    else:
        n = 1 if model == DLS_MODELS[1] else 2
        out['d'], out['p'], out['fit'], out['peaks'], out['rms'], out['bic'] = dls_fit_peaks(
            *args, n_peaks=n, gmin=0.05 if gmin is None else gmin, **lim)
    d, p = out['d'], out['p']
    mean = float(np.sum(d * p) / np.sum(p))
    out['mode'], out['mean'] = float(d[np.argmax(p)]), mean
    out['sd'] = float(np.sqrt(np.sum(p * (d - mean) ** 2) / np.sum(p)))
    return out


DLS_WEIGHTS = ("Intensity", "Volume", "Number")
DLS_CONVERSIONS = ("Mie", "Mie shell", "Rayleigh")
_DLS_POWER = {"Intensity": 0, "Volume": 3, "Number": 6}   # Rayleigh: I ~ N d^6, so volume ~ I / d^3 and number ~ I / d^6


def mie_intensity(d_nm, wavelength_nm, n_medium, n_particle, angle_deg):
    """Light scattered at `angle_deg` by one homogeneous non-absorbing sphere of each diameter in `d_nm`, perpendicular
    to the scattering plane (|S1|^2, vertically polarised laser, detector in the horizontal plane), in arbitrary units
    that are the same for every diameter. Mie theory, Bohren-Huffman coefficients (downward recurrence of the logarithmic
    derivative, upward recurrence of the Riccati-Bessel functions). Checked: Qext = Qsca (no absorption), the d^6 law
    of Rayleigh at small sizes and the forward sum Qext = 4/x^2 Re S(0)."""
    out = np.empty(len(d_nm))
    m = n_particle / n_medium
    mu = np.cos(np.radians(angle_deg))
    for i, d in enumerate(d_nm):
        x = np.pi * d * n_medium / wavelength_nm
        nmax = int(x + 4 * x ** (1 / 3) + 2)
        mx = m * x
        nmx = int(max(nmax, abs(mx))) + 15
        D = np.zeros(nmx + 1, complex)
        for n in range(nmx, 0, -1):
            D[n - 1] = n / mx - 1 / (D[n] + n / mx)
        psi0, psi1, chi0, chi1 = np.cos(x), np.sin(x), -np.sin(x), np.cos(x)
        xi1 = psi1 - 1j * chi1
        s1 = 0j
        pi_prev, pi_cur = 0.0, 1.0
        for n in range(1, nmax + 1):
            psi = (2 * n - 1) / x * psi1 - psi0
            chi = (2 * n - 1) / x * chi1 - chi0
            xi = psi - 1j * chi
            a = ((D[n] / m + n / x) * psi - psi1) / ((D[n] / m + n / x) * xi - xi1)
            b = ((m * D[n] + n / x) * psi - psi1) / ((m * D[n] + n / x) * xi - xi1)
            tau = n * mu * pi_cur - (n + 1) * pi_prev
            s1 += (2 * n + 1) / (n * (n + 1)) * (a * pi_cur + b * tau)
            pi_prev, pi_cur = pi_cur, ((2 * n + 1) * mu * pi_cur - (n + 1) * pi_prev) / n
            psi0, psi1, chi0, chi1 = psi1, psi, chi1, chi
            xi1 = psi1 - 1j * chi1
        out[i] = abs(s1) ** 2
    return out


def mie_shell_intensity(d_nm, wavelength_nm, n_medium, shell_dn, shell_nm, angle_deg):
    """Same as mie_intensity() for a vesicle: a shell of thickness `shell_nm` and index n_medium + `shell_dn` around a
    core filled with the medium (outer diameter d_nm; a diameter below twice the thickness is a solid sphere of the shell
    index). Coated-sphere coefficients of Bohren-Huffman as in LabSpectrumManager's vesicle model. Needs scipy."""
    from scipy.special import spherical_jn, spherical_yn

    def ric(n, z):
        j, yy = spherical_jn(n, z), spherical_yn(n, z)
        jp, yp = spherical_jn(n, z, derivative=True), spherical_yn(n, z, derivative=True)
        return z * j, j + z * jp, -z * yy, -(yy + z * yp)

    m2 = (n_medium + shell_dn) / n_medium
    mu = np.cos(np.radians(angle_deg))
    out = np.empty(len(d_nm))
    for i, d in enumerate(d_nm):
        k = 2 * np.pi * n_medium / wavelength_nm
        b = d / 2.0
        a = max(b - shell_nm, 1e-6)
        x, y = k * a, k * b
        n = np.arange(1, int(np.ceil(y + 4 * y ** (1 / 3) + 2)) + 1)
        p1x, d1x, _, _ = ric(n, x)
        p2x, d2x, c2x, dc2x = ric(n, m2 * x)
        py, dpy, cy, dcy = ric(n, y)
        p2y, d2y, c2y, dc2y = ric(n, m2 * y)
        xi, dxi = py - 1j * cy, dpy - 1j * dcy
        An = (m2 * p2x * d1x - d2x * p1x) / (m2 * c2x * d1x - dc2x * p1x)
        Bn = (m2 * p1x * d2x - p2x * d1x) / (m2 * dc2x * p1x - d1x * c2x)
        an = (py * (d2y - An * dc2y) - m2 * dpy * (p2y - An * c2y)) / (xi * (d2y - An * dc2y) - m2 * dxi * (p2y - An * c2y))
        bn = (m2 * py * (d2y - Bn * dc2y) - dpy * (p2y - Bn * c2y)) / (m2 * xi * (d2y - Bn * dc2y) - dxi * (p2y - Bn * c2y))
        s1, pi_prev, pi_cur = 0j, 0.0, 1.0
        for j in n:
            tau = j * mu * pi_cur - (j + 1) * pi_prev
            s1 += (2 * j + 1) / (j * (j + 1)) * (an[j - 1] * pi_cur + bn[j - 1] * tau)
            pi_prev, pi_cur = pi_cur, ((2 * j + 1) * mu * pi_cur - (j + 1) * pi_prev) / j
        out[i] = abs(s1) ** 2
    return out


def dls_per_particle(d, conversion, rec, wavelength, n_particle, shell=(4.5, 0.12)):
    """Intensity scattered by one particle of each diameter `d`: d^6 for 'Rayleigh', Mie theory for 'Mie' (needs the
    scattering angle and solvent index of `rec`, the laser wavelength and the refractive index of the particle)."""
    d = np.asarray(d, float)
    if conversion == "Rayleigh":
        return d ** 6
    if conversion == "Mie shell":
        return mie_shell_intensity(d, wavelength, rec['n'], shell[1], shell[0], rec['angle'])
    return mie_intensity(d, wavelength, rec['n'], n_particle, rec['angle'])


def dls_reweight(d, p, kind, per_particle=None):
    """Distribution `p` (weights per diameter `d`, intensity-weighted) converted to 'Intensity', 'Volume' or 'Number'
    weighting and normalised to sum 1. `per_particle` is the intensity one particle of each diameter scatters (from
    dls_per_particle): number ~ p / I1, volume ~ p d^3 / I1. Without it, the Rayleigh approximation I1 ~ d^6 is used
    (number ~ p / d^6, volume ~ p / d^3), which holds for particles much smaller than the wavelength in the medium
    (400 nm) and overweights the large ones above about 100 nm."""
    d = np.asarray(d, float)
    w = dls_reweight_weights(d, p, kind, d ** 6 if per_particle is None else per_particle)
    return w / w.sum()


def dls_reweight_peaks(peaks, kind, d=None, per_particle=None):
    """Log-normal peaks of an intensity distribution in the other weightings: new dicts {'median', 'sigma'?, 'fraction'}.
    Without `per_particle` (Rayleigh) it is analytic: multiplying a log-normal in ln d by d^-k gives a log-normal with
    the same sigma and the median divided by exp(k sigma^2), and the share multiplied by exp(-k mu + k^2 sigma^2 / 2).
    With `per_particle` (Mie, on the grid `d`) each peak is rebuilt on the grid, reweighted, and its median and share
    are read from it (the shape is no longer log-normal, so 'sigma' is the S.D. of ln d)."""
    if per_particle is None:
        k = _DLS_POWER[kind]
        out = []
        for pk in peaks:
            mu, s = np.log(pk['median']), pk['sigma']
            out.append({'median': float(np.exp(mu - k * s * s)), 'sigma': s,
                        'fraction': pk['fraction'] * float(np.exp(-k * mu + k * k * s * s / 2))})
        tot = sum(pk['fraction'] for pk in out)
        for pk in out:
            pk['fraction'] /= tot
        return out
    lnd = np.log(d)
    comps = []
    for pk in peaks:
        g = np.exp(-0.5 * ((lnd - np.log(pk['median'])) / pk['sigma']) ** 2)
        comps.append(pk['fraction'] * g / g.sum())
    ws = [dls_reweight_weights(d, c, kind, per_particle) for c in comps]
    total = sum(w.sum() for w in ws)
    out = []
    for w in ws:
        cum = (np.cumsum(w) - w / 2) / w.sum()          # cumulative at the bin centres, not at their upper edges
        med = float(np.exp(np.interp(0.5, cum, lnd)))
        sd = float(np.sqrt(np.sum(w * (lnd - np.sum(w * lnd) / w.sum()) ** 2) / w.sum()))
        out.append({'median': med, 'sigma': sd, 'fraction': float(w.sum() / total)})
    return out


def dls_reweight_weights(d, p, kind, per_particle):
    """Unnormalised weights of `p` in the weighting `kind` (the numerator of dls_reweight)."""
    d = np.asarray(d, float)
    return np.asarray(p, float) * {"Intensity": 1.0, "Volume": d ** 3, "Number": 1.0}[kind] / (
        1.0 if kind == "Intensity" else np.asarray(per_particle, float))


def dls_floor(p, fraction):
    """`p` with the bins below `fraction` of its maximum set to zero. The regularised inversion leaves a few percent of
    intensity at the edge of the size grid; dividing by the light one particle scatters (d^6 for small sizes) turns
    it into most of the particles, so it is cut before the conversion to volume or number."""
    p = np.asarray(p, float)
    return np.where(p >= fraction * p.max(), p, 0.0)


def dls_stats(d, p):
    """Mean, mode and standard deviation (nm) of a distribution `p` on the diameters `d`."""
    mean = float(np.sum(d * p) / np.sum(p))
    return mean, float(d[np.argmax(p)]), float(np.sqrt(np.sum(p * (d - mean) ** 2) / np.sum(p)))


def dls_record(d):
    """The input of dls_analyse() for a sample dict of load_nsz(): delays in seconds and the measurement conditions
    (angle, solvent index, viscosity, temperature) read from the file. None if the file lacks any of them or the delays
    are not matched to the channels."""
    p = d['params']
    keys = ('Angle_deg', 'n_solvent', 'Visc_mPas', 'Temp_C')
    if not d.get('tau_aligned') or any(p.get(k) is None for k in keys) or len(d['tau']) < 8:
        return None
    rec = {'tau': d['tau'] * 1e-6, 'acf': d['acf'], 'angle': p['Angle_deg'], 'n': p['n_solvent'],
           'visc': p['Visc_mPas'], 'T': p['Temp_C']}
    if p.get('Range_from'):
        rec['range'] = (p['Range_from'] - 1, p['Range_to'])      # 0-based [first, end) channel indices
    return rec


# ── Main application ──────────────────────────────────────────────────────────

class DLSManager:
    def __init__(self, root):
        self.root = root
        self.root.title("DLS Manager — Horiba SZ-100")
        self.root.geometry("1380x840")
        if psk is None and not _psk_config().get('declined'):
            self.root.after(400, lambda: self._offer_plotstylekit(PSK_PROBLEM))
        self.root.configure(bg="#F1F5F9")

        self.samples = {}   # path → data dict
        self.selected = []  # list of paths currently shown
        self._dls_cache = {}   # (path, model, wavelength, alpha, g1 min) -> dls_analyse() result

        self._build_menu()
        self._build_ui()
        self._setup_dnd()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_menu(self):
        bar = tk.Menu(self.root)
        m = tk.Menu(bar, tearoff=0)
        m.add_command(label="Open .nsz files...", command=self._open_files)
        m.add_command(label="Open folder...", command=self._open_folder)
        m.add_command(label="Export Excel...", command=self._export_xlsx)
        m.add_separator()
        m.add_command(label="Save Figure Image (PNG, PDF, SVG)...", command=self.save_figure_image)
        m.add_command(label="Save Figure (pickle)...", command=self.save_figure_pickle)
        m.add_command(label="Edit Figure...", command=self.open_figure_editor)
        m.add_command(label="PlotStyleKit...", command=self.plotstylekit_dialog)
        m.add_separator()
        m.add_command(label="Exit", command=self.root.destroy)
        bar.add_cascade(label="File", menu=m)
        self.root.config(menu=bar)

    def _build_ui(self):
        # ── Left sidebar ──────────────────────────────────────────────────────
        sidebar = tk.Frame(self.root, bg="#1E293B", width=260)
        sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=0, pady=0)
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="DLS Manager", font=("Helvetica", 15, "bold"),
                 fg="white", bg="#1E293B").pack(pady=(18,4))
        tk.Label(sidebar, text="Horiba SZ-100 DLS", font=("Helvetica", 9),
                 fg="#94A3B8", bg="#1E293B").pack(pady=(0,14))

        # Buttons
        btn_frame = tk.Frame(sidebar, bg="#1E293B")
        btn_frame.pack(fill=tk.X, padx=10, pady=(0,10))

        for text, cmd in [
            ("＋  Open .nsz files", self._open_files),
            ("📁  Open folder",  self._open_folder),
            ("🗑  Remove selected",   self._remove_selected),
            ("💾  Export Excel",  self._export_xlsx),
        ]:
            b = tk.Button(btn_frame, text=text, command=cmd,
                          bg="#334155", fg="white", relief=tk.FLAT,
                          font=("Helvetica", 10), anchor="w",
                          padx=10, pady=6, cursor="hand2",
                          activebackground="#475569", activeforeground="white")
            b.pack(fill=tk.X, pady=2)

        # Drag & drop zone
        drop_lbl = tk.Label(sidebar,
            text="⬇ Drop .nsz files here",
            font=("Helvetica", 9), fg="#64748B", bg="#1E293B",
            pady=10)
        drop_lbl.pack(fill=tk.X, padx=10)

        ttk.Separator(sidebar, orient="horizontal").pack(fill=tk.X, padx=10, pady=8)

        # Sample list
        tk.Label(sidebar, text="LOADED SAMPLES", font=("Helvetica", 8, "bold"),
                 fg="#64748B", bg="#1E293B").pack(anchor="w", padx=12)

        list_frame = tk.Frame(sidebar, bg="#1E293B")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        self.listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED,
                                  bg="#0F172A", fg="white",
                                  selectbackground="#2563EB",
                                  font=("Helvetica", 9),
                                  borderwidth=0, highlightthickness=0,
                                  activestyle="none", exportselection=False)
        scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                               command=self.listbox.yview)
        self.listbox.config(yscrollcommand=scroll.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        # ── Main area ─────────────────────────────────────────────────────────
        main = tk.Frame(self.root, bg="#F1F5F9")
        main.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Tabs
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook", background="#F1F5F9", borderwidth=0)
        style.configure("TNotebook.Tab", padding=[14,7], font=("Helvetica", 10))
        style.map("TNotebook.Tab",
                  background=[("selected","white"),("!selected","#E2E8F0")],
                  foreground=[("selected","#1E293B"),("!selected","#64748B")])

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Tab 1: ACF
        self.tab_acf = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.tab_acf, text="  Autocorrelogram  ")

        # Tab 2: Size distribution
        self.tab_size = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.tab_size, text="  Distribution  ")

        # Tab 3: Parameters
        self.tab_params = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.tab_params, text="  Parameters  ")

        # Tab 4: DLS analysis of the stored correlogram
        self.tab_dls = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.tab_dls, text="  DLS analysis  ")

        self._build_acf_tab()
        self._build_size_tab()
        self._build_params_tab()
        self._build_dls_tab()
        # the inversions take a moment: they run only while their tab is shown
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self._refresh_dls())

    def _build_acf_tab(self):
        ctrl = tk.Frame(self.tab_acf, bg="white", pady=4)
        ctrl.pack(fill=tk.X, padx=14)

        tk.Label(ctrl, text="X scale:", bg="white", font=("Helvetica",9)).pack(side=tk.LEFT)
        self.acf_xscale = tk.StringVar(value="log")
        for val, lbl in [("log","Log"),("linear","Linear")]:
            tk.Radiobutton(ctrl, text=lbl, variable=self.acf_xscale, value=val,
                           bg="white", font=("Helvetica",9),
                           command=self._plot_acf).pack(side=tk.LEFT, padx=4)

        self.show_fit = tk.BooleanVar(value=True)
        tk.Checkbutton(ctrl, text="Show ideal fit", variable=self.show_fit,
                       bg="white", font=("Helvetica",9),
                       command=self._plot_acf).pack(side=tk.LEFT, padx=12)

        self.show_resid = tk.BooleanVar(value=False)
        tk.Checkbutton(ctrl, text="Show residuals", variable=self.show_resid,
                       bg="white", font=("Helvetica",9),
                       command=self._plot_acf).pack(side=tk.LEFT, padx=4)

        self.fig_acf = Figure(figsize=(9,5.5), dpi=100, facecolor="white")
        self.canvas_acf = FigureCanvasTkAgg(self.fig_acf, self.tab_acf)
        NavigationToolbar2Tk(self.canvas_acf, self.tab_acf).pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas_acf.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

    def _build_size_tab(self):
        ctrl = tk.Frame(self.tab_size, bg="white", pady=4)
        ctrl.pack(fill=tk.X, padx=14)

        tk.Label(ctrl, text="X scale:", bg="white", font=("Helvetica",9)).pack(side=tk.LEFT)
        self.size_xscale = tk.StringVar(value="log")
        for val, lbl in [("log","Log"),("linear","Linear")]:
            tk.Radiobutton(ctrl, text=lbl, variable=self.size_xscale, value=val,
                           bg="white", font=("Helvetica",9),
                           command=self._plot_size).pack(side=tk.LEFT, padx=4)

        ttk.Separator(ctrl, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        tk.Label(ctrl, text="Representation:", bg="white", font=("Helvetica",9)).pack(side=tk.LEFT)
        self.size_repr = tk.StringVar(value="intensity")
        for val, lbl in [("intensity","Intensity"),("number","Number")]:
            tk.Radiobutton(ctrl, text=lbl, variable=self.size_repr, value=val,
                           bg="white", font=("Helvetica",9),
                           command=self._plot_size).pack(side=tk.LEFT, padx=4)

        ttk.Separator(ctrl, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        self.show_cumul = tk.BooleanVar(value=True)
        tk.Checkbutton(ctrl, text="Show cumulative", variable=self.show_cumul,
                       bg="white", font=("Helvetica",9),
                       command=self._plot_size).pack(side=tk.LEFT, padx=4)

        self.fig_size = Figure(figsize=(9,5.5), dpi=100, facecolor="white")
        self.canvas_size = FigureCanvasTkAgg(self.fig_size, self.tab_size)
        NavigationToolbar2Tk(self.canvas_size, self.tab_size).pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas_size.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

    def _build_params_tab(self):
        frame = tk.Frame(self.tab_params, bg="white")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        cols = ("Parameter",
                "Sample 1","Sample 2","Sample 3",
                "Sample 4","Sample 5")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings",
                                 selectmode="browse")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.heading("Parameter", text="Parameter")
        self.tree.column("Parameter", width=160, anchor="w")
        for c in cols[1:]:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=180, anchor="center")

        style = ttk.Style()
        style.configure("Treeview", font=("Helvetica",9), rowheight=22)
        style.configure("Treeview.Heading", font=("Helvetica",9,"bold"))
        style.map("Treeview",
                  background=[("selected","#DBEAFE")],
                  foreground=[("selected","#1E3A5F")])

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Alternating rows
        self.tree.tag_configure("odd",  background="#F8FAFC")
        self.tree.tag_configure("even", background="white")

    def _build_dls_tab(self):
        ctrl = tk.Frame(self.tab_dls, bg="white", pady=4)
        ctrl.pack(fill=tk.X, padx=14)
        self._tips = []
        tk.Label(ctrl, text="Model:", bg="white", font=("Helvetica", 9)).pack(side=tk.LEFT)
        self.dls_model = tk.StringVar(value=DLS_MODELS[0])
        cmb = ttk.Combobox(ctrl, textvariable=self.dls_model, values=DLS_MODELS, state="readonly", width=30)
        cmb.pack(side=tk.LEFT, padx=(4, 0))
        self._tips.append(ToolTip(cmb, DLS_TIPS['model']))
        self.dls_vars = {}
        for text, key, val in (("Wavelength (nm)", 'wl', '532'), ("Smoothing α", 'alpha', '1.0'),
                               ("g1 min", 'gmin', '0.1')):
            lab = tk.Label(ctrl, text=text, bg="white", font=("Helvetica", 9))
            lab.pack(side=tk.LEFT, padx=(12, 2))
            self.dls_vars[key] = tk.StringVar(value=val)
            ent = tk.Entry(ctrl, textvariable=self.dls_vars[key], width=6)
            ent.pack(side=tk.LEFT)
            ent.bind("<Return>", lambda e: self._plot_dls())
            self._tips.extend(ToolTip(w, DLS_TIPS[key]) for w in (lab, ent))
        # the usual g1 cut-off depends on the model (0.1 regularised, 0.05 peaks)
        self.dls_model.trace_add("write", self._on_dls_model)
        ctrl2 = tk.Frame(self.tab_dls, bg="white")
        ctrl2.pack(fill=tk.X, padx=14)
        tk.Label(ctrl2, text="Weighting:", bg="white", font=("Helvetica", 9)).pack(side=tk.LEFT)
        self.dls_weight = tk.StringVar(value=DLS_WEIGHTS[0])
        for w in DLS_WEIGHTS:
            rb = tk.Radiobutton(ctrl2, text=w, variable=self.dls_weight, value=w, bg="white", font=("Helvetica", 9),
                                command=self._plot_dls)
            rb.pack(side=tk.LEFT, padx=2)
            self._tips.append(ToolTip(rb, DLS_TIPS['weight']))
        self.dls_conv = tk.StringVar(value=DLS_CONVERSIONS[0])
        cv = ttk.Combobox(ctrl2, textvariable=self.dls_conv, values=DLS_CONVERSIONS, state="readonly", width=9)
        cv.pack(side=tk.LEFT, padx=(8, 0))
        cv.bind("<<ComboboxSelected>>", lambda e: self._plot_dls())
        self.dls_range = tk.BooleanVar(value=True)
        rb = tk.Checkbutton(ctrl2, text="Instrument range", variable=self.dls_range, bg="white", font=("Helvetica", 9),
                            command=self._plot_dls)
        rb.pack(side=tk.LEFT, padx=(12, 0))
        self._tips.append(ToolTip(rb, DLS_TIPS['range']))
        lab = tk.Label(ctrl2, text="Tail cut (% of max)", bg="white", font=("Helvetica", 9))
        lab.pack(side=tk.LEFT, padx=(8, 2))
        self.dls_vars['tail'] = tk.StringVar(value='5')
        ent = tk.Entry(ctrl2, textvariable=self.dls_vars['tail'], width=4)
        ent.pack(side=tk.LEFT)
        ent.bind("<Return>", lambda e: self._plot_dls())
        self._tips.extend(ToolTip(w, DLS_TIPS['tail']) for w in (lab, ent))
        lab = tk.Label(ctrl2, text="Particle n", bg="white", font=("Helvetica", 9))
        lab.pack(side=tk.LEFT, padx=(8, 2))
        self.dls_vars['np'] = tk.StringVar(value='')    # empty = the index stored in the file (1.59 if absent)
        ent = tk.Entry(ctrl2, textvariable=self.dls_vars['np'], width=6)
        ent.pack(side=tk.LEFT)
        ent.bind("<Return>", lambda e: self._plot_dls())
        self._tips.extend(ToolTip(w, DLS_TIPS['conv']) for w in (cv, lab, ent))
        ctrl3 = tk.Frame(self.tab_dls, bg="white")
        ctrl3.pack(fill=tk.X, padx=14)
        tk.Label(ctrl3, text="Vesicle shell (Mie shell):", bg="white", font=("Helvetica", 9)).pack(side=tk.LEFT)
        for text, key, val in (("thickness (nm)", 'shell_t', '4.5'), ("index excess", 'shell_dn', '0.12')):
            lab = tk.Label(ctrl3, text=text, bg="white", font=("Helvetica", 9))
            lab.pack(side=tk.LEFT, padx=(8, 2))
            self.dls_vars[key] = tk.StringVar(value=val)
            ent = tk.Entry(ctrl3, textvariable=self.dls_vars[key], width=6)
            ent.pack(side=tk.LEFT)
            ent.bind("<Return>", lambda e: self._plot_dls())
            self._tips.extend(ToolTip(w, DLS_TIPS['shell']) for w in (lab, ent))
        btn = tk.Button(ctrl2, text="Compute", command=self._plot_dls, bg="#FEF9C3", font=("Helvetica", 9))
        btn.pack(side=tk.LEFT, padx=12)

        self.dls_text = tk.Text(self.tab_dls, height=9, font=("Consolas", 9), bg="#F8FAFC", relief=tk.FLAT,
                                wrap=tk.WORD, state=tk.DISABLED)
        self.dls_text.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 6))
        self.fig_dls = Figure(figsize=(9, 4.5), dpi=100, facecolor="white")
        self.canvas_dls = FigureCanvasTkAgg(self.fig_dls, self.tab_dls)
        NavigationToolbar2Tk(self.canvas_dls, self.tab_dls).pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas_dls.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

    # ── Drag & Drop ───────────────────────────────────────────────────────────

    def _setup_dnd(self):
        if DND_AVAILABLE:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self._on_drop)
        # fallback: right click on the list opens the file dialog
        self.listbox.bind("<Button-3>", lambda e: self._open_files())

    def _on_drop(self, event):
        raw = event.data
        # tkinterdnd2 returns space-separated paths, with {} around those containing spaces
        paths = self.root.tk.splitlist(raw)
        nsz_files = []
        for p in paths:
            if os.path.isfile(p) and p.lower().endswith('.nsz'):
                nsz_files.append(p)
            elif os.path.isdir(p):
                for f in os.listdir(p):
                    if f.lower().endswith('.nsz'):
                        nsz_files.append(os.path.join(p, f))
        self._load_files(nsz_files)

    # ── File loading ──────────────────────────────────────────────────────────

    def _open_files(self):
        files = filedialog.askopenfilenames(
            title="Open NSZ files",
            filetypes=[("Horiba NSZ", "*.nsz"), ("All Files", "*.*")])
        self._load_files(files)

    def _open_folder(self):
        folder = filedialog.askdirectory(title="Open folder with NSZ files")
        if folder:
            files = [os.path.join(folder, f)
                     for f in sorted(os.listdir(folder))
                     if f.lower().endswith('.nsz')]
            self._load_files(files)

    def _load_files(self, paths):
        new = 0
        errors = []
        for path in paths:
            path = os.path.normpath(path)
            if path in self.samples:
                continue
            try:
                data = load_nsz(path)
                self.samples[path] = data
                short = os.path.basename(path)
                self.listbox.insert(tk.END, short)
                new += 1
            except Exception as e:
                errors.append(f"{os.path.basename(path)}: {e}")

        if errors:
            messagebox.showerror("Loading error",
                                 "\n".join(errors[:5]))
        if new:
            # Automatically select the new ones
            n = self.listbox.size()
            self.listbox.selection_clear(0, tk.END)
            for i in range(n - new, n):
                self.listbox.selection_set(i)
            self._on_select()

    def _remove_selected(self):
        sel = list(self.listbox.curselection())
        paths = list(self.samples.keys())
        for i in reversed(sel):
            path = paths[i]
            del self.samples[path]
            self.listbox.delete(i)
        self.selected = []
        self._refresh_all()

    # ── Selection ─────────────────────────────────────────────────────────────

    def _on_select(self, event=None):
        indices = self.listbox.curselection()
        paths = list(self.samples.keys())
        self.selected = [paths[i] for i in indices if i < len(paths)]
        self._refresh_all()

    def _refresh_all(self):
        self._plot_acf()
        self._plot_size()
        self._update_params_table()
        self._refresh_dls()

    def _refresh_dls(self):
        """Redraw the DLS tab, but only while it is the visible one."""
        if self.notebook.select() == str(self.tab_dls):
            self._plot_dls()

    # ── Plot style (PlotStyleKit) ─────────────────────────────────────────────

    def _style_figure(self, fig):
        """Apply the shared Origin style to every axes of `fig`, then enlarge the fonts for the on-screen
        pane (the 8-9 pt paper sizes look tiny when the plot fills the tab) and lay the figure out."""
        if psk is None:
            for ax in fig.axes:
                ax.grid(True, which='both', alpha=0.3, linestyle='--')
            fig.tight_layout()
            return
        for ax in fig.axes:
            psk.style(ax, None, set_size=False)
            ax.title.set_fontsize(11)
            ax.xaxis.label.set_fontsize(10)
            ax.yaxis.label.set_fontsize(10)
            ax.tick_params(axis='both', labelsize=9)
            leg = ax.get_legend()
            if leg is not None:
                for t in leg.get_texts():
                    t.set_fontsize(9)
        fig._editor_font_family = psk.origin_style().font_origin()
        fig.tight_layout()

    # ── ACF plot ──────────────────────────────────────────────────────────────

    def _plot_acf(self):
        self.fig_acf.clear()

        if self.show_resid.get() and self.selected:
            gs = self.fig_acf.add_gridspec(2, 1, height_ratios=[3,1], hspace=0.08)
            ax  = self.fig_acf.add_subplot(gs[0])
            axr = self.fig_acf.add_subplot(gs[1], sharex=ax)
        else:
            ax  = self.fig_acf.add_subplot(111)
            axr = None

        xscale = self.acf_xscale.get()
        palette = sample_colors()

        for idx, path in enumerate(self.selected[:MAX_SAMPLES_SHOWN]):
            d = self.samples[path]
            c = palette[idx % len(palette)]
            lbl = self._short_label(d['name'])

            tau = d['tau']
            acf = d['acf']

            # Keep sensible values only
            mask = (tau > 0) & (acf >= -0.05) & (acf <= 1.05)
            if mask.sum() < 2:
                continue

            ax.plot(tau[mask], acf[mask], 'o', color=c, markersize=3.5,
                    linewidth=0, label=lbl)

            # Ideal fit — aligned to tau[fit_offset:]
            off = d.get('fit_offset', 6)
            if self.show_fit.get() and len(d['fit']) > 2:
                tau_fit = tau[off:off + len(d['fit'])]
                fit_vals = d['fit'][:len(tau_fit)]
                ax.plot(tau_fit, fit_vals, '-', color=c, linewidth=1.5,
                        alpha=0.85)

            # Residuals — same offset as the fit
            if axr is not None and len(d['resid']) > 2:
                tau_res = tau[off:off + len(d['resid'])]
                res_vals = d['resid'][:len(tau_res)]
                axr.plot(tau_res, res_vals, 'o', color=c, markersize=2.5,
                         linewidth=1, alpha=0.8)

        ax.set_xscale(xscale)
        ax.set_xlabel("Delay time τ (µs)", fontsize=10)
        # mathtext because Arial has no subscript glyphs; the stream holds g1 squared (see CLAUDE.md)
        ax.set_ylabel(r"$g_2(\tau) - 1$ normalized", fontsize=10)
        ax.set_ylim(-0.05, 1.1)
        ax.set_title("Autocorrelation function (ACF)", fontsize=11, fontweight='bold')
        ax.axhline(0, color='gray', linewidth=0.7, linestyle='--')
        if self.selected:
            ax.legend(fontsize=8, loc='upper right', framealpha=0.9)

        if axr is not None:
            axr.set_xscale(xscale)
            axr.set_xlabel("Delay time τ (µs)", fontsize=10)
            axr.set_ylabel("Residuals", fontsize=9)
            axr.axhline(0, color='gray', linewidth=0.8)
            plt.setp(ax.get_xticklabels(), visible=False)
            ax.set_xlabel("")

        if not self.selected:
            ax.text(0.5, 0.5, "No sample selected\n\nLoad .nsz files and select them in the list",
                    ha='center', va='center', transform=ax.transAxes,
                    fontsize=12, color='#94A3B8')

        self._style_figure(self.fig_acf)
        self.canvas_acf.draw()

    # ── Distribution plot ─────────────────────────────────────────────────────

    @staticmethod
    def _to_number_dist(sz, intens):
        """Convert an intensity distribution to a number distribution (Rayleigh approximation: N ∝ I/d⁶)."""
        d6 = np.where(sz > 0, sz ** 6, np.nan)
        num = np.where(sz > 0, intens / d6, 0.0)
        total = num.sum()
        if total > 0:
            num = num / total * 100.0
        cumul = np.cumsum(num)
        if cumul[-1] > 0:
            cumul = cumul / cumul[-1] * 100.0
        return num, cumul

    def _refresh_size(self):
        self._plot_size()
        self._update_params_table()

    def _plot_size(self):
        self.fig_size.clear()

        by_number = self.size_repr.get() == "number"
        show_c = self.show_cumul.get()
        if show_c and self.selected:
            gs = self.fig_size.add_gridspec(1, 2, wspace=0.3)
            ax  = self.fig_size.add_subplot(gs[0])
            axc = self.fig_size.add_subplot(gs[1])
        else:
            ax  = self.fig_size.add_subplot(111)
            axc = None

        xscale = self.size_xscale.get()
        palette = sample_colors()

        for idx, path in enumerate(self.selected[:MAX_SAMPLES_SHOWN]):
            d = self.samples[path]
            c = palette[idx % len(palette)]
            lbl = self._short_label(d['name'])

            sz    = d['sz']
            mask  = sz > 0
            if mask.sum() < 2:
                continue

            if by_number:
                freq, cumul = self._to_number_dist(sz, d['intens'])
            else:
                freq  = d['intens']
                cumul = d['cumul']

            ax.plot(sz[mask], freq[mask], 'o-', color=c, markersize=3,
                    linewidth=1.5, label=lbl)
            ax.fill_between(sz[mask], freq[mask], alpha=0.08, color=c)

            p = d['params']
            if not by_number and p.get('D50_nm'):
                ax.axvline(p['D50_nm'], color=c, linewidth=0.8,
                           linestyle=':', alpha=0.7)

            if axc is not None:
                axc.plot(sz[mask], cumul[mask], 'o-', color=c, markersize=3,
                         linewidth=1.5, label=lbl)
                if not by_number:
                    for pct in [10, 50, 90]:
                        val = p.get(f'D{pct}_nm')
                        if val:
                            axc.axvline(val, color=c, linewidth=0.7,
                                        linestyle='--', alpha=0.6)

        repr_lbl = "number" if by_number else "intensity"
        ax.margins(y=0.18)                  # headroom so the legend does not cover the top of the peak
        top = ax.get_ylim()[1]
        ax.set_ylim(-0.02 * top, top)       # distributions are >= 0: no empty space below the baseline
        ax.set_xscale(xscale)
        ax.set_xlabel("Diameter (nm)", fontsize=10)
        ax.set_ylabel(f"{'Number' if by_number else 'Intensity'} (%)", fontsize=10)
        ax.set_title(f"Size distribution by {repr_lbl}", fontsize=11, fontweight='bold')
        if self.selected:
            ax.legend(fontsize=8, loc='upper right', framealpha=0.9)

        if axc is not None:
            axc.set_xscale(xscale)
            axc.set_xlabel("Diameter (nm)", fontsize=10)
            axc.set_ylabel("Cumulative (%)", fontsize=10)
            axc.set_title(f"Cumulative distribution by {repr_lbl}", fontsize=11, fontweight='bold')
            axc.set_ylim(-2, 104)
            axc.axhline(10, color='gray', linewidth=0.5, linestyle=':')
            axc.axhline(50, color='gray', linewidth=0.5, linestyle=':')
            axc.axhline(90, color='gray', linewidth=0.5, linestyle=':')
            if self.selected:
                axc.legend(fontsize=8, loc='upper left', framealpha=0.9)

        if not self.selected:
            ax.text(0.5, 0.5, "No sample selected\n\nLoad .nsz files and select them in the list",
                    ha='center', va='center', transform=ax.transAxes,
                    fontsize=12, color='#94A3B8')

        self._style_figure(self.fig_size)
        self.canvas_size.draw()

    # ── DLS analysis ──────────────────────────────────────────────────────────

    def _on_dls_model(self, *_):
        self.dls_vars['gmin'].set('0.1' if self.dls_model.get() == DLS_MODELS[0] else '0.05')
        self._plot_dls()

    def _dls_result(self, path, model, wl, alpha, gmin):
        use_range = self.dls_range.get()
        key = (path, model, wl, alpha, gmin, use_range)
        if key not in self._dls_cache:
            self._dls_cache[key] = dls_analyse(dls_record(self.samples[path]), model, wl, alpha, gmin, use_range)
        return self._dls_cache[key]

    def _set_dls_text(self, lines):
        self.dls_text.config(state=tk.NORMAL)
        self.dls_text.delete("1.0", tk.END)
        self.dls_text.insert(tk.END, "\n".join(lines))
        self.dls_text.config(state=tk.DISABLED)

    def _plot_dls(self):
        """Cumulants and intensity-weighted size distribution of the selected samples from their correlograms, with
        the instrument's own values next to them."""
        self.fig_dls.clear()
        ax1 = self.fig_dls.add_subplot(121)
        ax2 = self.fig_dls.add_subplot(122)
        lines = []
        try:
            wl, alpha, gmin = (float(self.dls_vars[k].get()) for k in ('wl', 'alpha', 'gmin'))
            if wl <= 0 or alpha < 0 or not 0 < gmin < 1:
                raise ValueError
        except ValueError:
            wl = None
            lines.append("Invalid settings: wavelength > 0, smoothing >= 0 and 0 < g1 min < 1 are required.")
        try:
            import scipy.optimize  # noqa: F401  (nnls and least_squares)
        except ImportError:
            wl = None
            lines.append("The inversion needs scipy (pip install scipy).")
        model = self.dls_model.get()
        weight = self.dls_weight.get()
        conv = self.dls_conv.get()
        palette = sample_colors()
        if wl is not None and self.selected:
            self.root.config(cursor="watch")
            self.root.update_idletasks()
        try:
            for idx, path in enumerate(self.selected[:MAX_SAMPLES_SHOWN] if wl is not None else []):
                d = self.samples[path]
                c = palette[idx % len(palette)]
                name = self._short_label(d['name'])
                rec = dls_record(d)
                if rec is None:
                    lines.append(f"{name}: the file has no correlogram or measurement conditions to analyse.")
                    continue
                try:
                    r = self._dls_result(path, model, wl, alpha, gmin)
                except Exception as e:
                    lines.append(f"{name}: inversion failed ({e}).")
                    continue
                g1 = np.sqrt(np.clip(rec['acf'], 0, None))
                if self.dls_range.get() and rec.get('range'):
                    lo, hi = rec['range']
                    out_of = np.ones(len(g1), bool)
                    out_of[lo:hi] = False
                    ax1.semilogx(rec['tau'][out_of] * 1e6, g1[out_of], 'o', ms=3, mfc='none', color=c, alpha=0.5)
                    ax1.semilogx(rec['tau'][~out_of] * 1e6, g1[~out_of], 'o', ms=3, color=c, label=name)
                else:
                    ax1.semilogx(rec['tau'] * 1e6, g1, 'o', ms=3, color=c, label=name)
                ax1.semilogx(rec['tau'] * 1e6, r['fit'], '-', lw=1.3, color=c)
                i1 = self._dls_i1(d, rec, r['d'], conv, wl) if weight != DLS_WEIGHTS[0] else None
                pw = self._weighted(r, weight, i1)
                ax2.semilogx(r['d'], pw / pw.max(), '-', lw=1.6, color=c, label=name)
                if r['z'] is not None:
                    ax2.axvline(r['z'], color=c, ls=':', lw=1)
                lines.extend(self._dls_summary(path, name, r, model, wl, alpha, gmin, weight, i1, conv))
        finally:
            self.root.config(cursor="")
        ax1.set_xlabel("Delay time τ (µs)")
        ax1.set_ylabel("$g_1$ = √ACF")
        ax1.set_title("Correlogram and fit", fontweight='bold')
        ax2.set_xlabel("Diameter (nm)")
        ax2.set_ylabel(f"{weight}-weighted (normalised)")
        ax2.set_title(f"Size distribution by {weight.lower()}", fontweight='bold')
        if self.selected and ax1.get_legend_handles_labels()[0]:
            ax1.legend(fontsize=8, loc='upper right', framealpha=0.9)
        if not self.selected:
            ax1.text(0.5, 0.5, "No sample selected\n\nLoad .nsz files and select them in the list",
                     ha='center', va='center', transform=ax1.transAxes, fontsize=12, color='#94A3B8')
        self._set_dls_text(lines)
        self._style_figure(self.fig_dls)
        self.canvas_dls.draw()

    def _tail_fraction(self):
        try:
            return min(max(float(self.dls_vars['tail'].get()), 0.0), 50.0) / 100.0
        except ValueError:
            return 0.05

    def _weighted(self, r, weight, i1):
        """Distribution `r` of dls_analyse() in the chosen weighting (the tail cut applies to the free model)."""
        p = r['p'] if r['peaks'] is not None or weight == DLS_WEIGHTS[0] else dls_floor(r['p'], self._tail_fraction())
        return dls_reweight(r['d'], p, weight, i1)

    def _particle_index(self, d):
        """Refractive index of the particle: the value typed in the tab, else the one in the file, else 1.59."""
        try:
            v = float(self.dls_vars['np'].get())
            if v > 1.0:
                return v
        except ValueError:
            pass
        return d['params'].get('n_particle') or 1.59

    def _shell(self):
        """(thickness nm, index excess) typed for the vesicle shell; the defaults of LabSpectrumManager if invalid."""
        try:
            t, dn = float(self.dls_vars['shell_t'].get()), float(self.dls_vars['shell_dn'].get())
            if t > 0 and 0 < dn < 1:
                return t, dn
        except ValueError:
            pass
        return 4.5, 0.12

    def _dls_i1(self, d, rec, grid, conv, wl):
        """Light one particle scatters at each diameter of `grid` (None for Rayleigh: dls_reweight then uses d^6)."""
        if conv == "Rayleigh":
            return None
        key = ('i1', d['path'], conv, wl, self._particle_index(d), self._shell(), len(grid))
        if key not in self._dls_cache:
            self._dls_cache[key] = dls_per_particle(grid, conv, rec, wl, self._particle_index(d), self._shell())
        return self._dls_cache[key]

    def _dls_summary(self, path, name, r, model, wl, alpha, gmin, weight=DLS_WEIGHTS[0], i1=None, conv="Rayleigh"):
        """Text lines for one sample: cumulants and distribution, each with the instrument's value."""
        p = self.samples[path]['params']

        def inst(key, fmt):
            return fmt.format(p[key]) if p.get(key) is not None else "n/a"

        rec = dls_record(self.samples[path])
        out = [f"{name}" + (f"  [instrument range: channels {rec['range'][0] + 1}-{rec['range'][1]}, "
                            f"{rec['tau'][rec['range'][0]] * 1e6:.3g}-{rec['tau'][rec['range'][1] - 1] * 1e6:.4g} µs]"
                            if self.dls_range.get() and rec.get('range') else "")]
        if r['z'] is not None:
            diff = f" ({100 * (r['z'] / p['ZAvg_nm'] - 1):+.1f}%)" if p.get('ZAvg_nm') else ""
            out.append(f"  Z-average {r['z']:.1f} nm, PI {r['pi']:.3f} (cumulants, g1 > 0.3)  |  instrument "
                       f"{inst('ZAvg_nm', '{:.1f}')} nm, PI {inst('PDI', '{:.3f}')}{diff}")
        w = weight.lower()
        by_int = weight == DLS_WEIGHTS[0]
        how = {"Mie": f"Mie, particle n {self._particle_index(self.samples[path]):g}",
               "Mie shell": f"Mie shell, t {self._shell()[0]:g} nm, index excess {self._shell()[1]:g}"}.get(conv, "Rayleigh (d^6)")
        if r['peaks'] is None:
            mean, mode, sd = dls_stats(r['d'], self._weighted(r, weight, i1))
            out.append(f"  Regularised distribution by {w} (α = {alpha:g}{'' if by_int else f', tail cut {100 * self._tail_fraction():g}%'}): mean {mean:.0f} nm, mode {mode:.0f} nm, "
                       f"S.D. {sd:.0f} nm" + (f"  |  instrument mean {inst('Mean1_nm', '{:.0f}')}, "
                       f"mode {inst('Mode_nm', '{:.0f}')}, S.D. {inst('SD1_nm', '{:.0f}')} nm" if by_int else
                       f"  ({how} conversion from the intensity distribution)"))
        else:
            for i, pk in enumerate(dls_reweight_peaks(r['peaks'], weight, r['d'], i1), 1):
                out.append(f"  Peak {i}: median {pk['median']:.0f} nm, σ(ln d) {pk['sigma']:.2f}, "
                           f"{100 * pk['fraction']:.0f}% of the {'scattered intensity' if by_int else w}")
            out.append(f"  rms of g1 {r['rms']:.4f}, BIC {r['bic']:.1f} (lower is better)" +
                       (f"  |  instrument mean {inst('Mean1_nm', '{:.0f}')}, mode {inst('Mode_nm', '{:.0f}')} nm"
                        if by_int else f"  (peaks converted from intensity by {how})"))
            if len(r['peaks']) == 2:
                r1 = self._dls_result(path, DLS_MODELS[1], wl, alpha, gmin)
                verdict = ('two peaks are supported' if r['bic'] < r1['bic'] - 6 else
                           'one peak is enough' if r['bic'] > r1['bic'] + 6 else 'not significant')
                out.append(f"  One peak: rms {r1['rms']:.4f}, BIC {r1['bic']:.1f}; difference "
                           f"{r1['bic'] - r['bic']:+.1f}: {verdict}")
        if not by_int:
            edge = r['peaks'] is None and (r['d'][np.argmax(r['p'])] >= r['d'][-1] * 0.99 or r['d'][np.argmax(r['p'])] <= r['d'][0] * 1.01)
            wide = r['peaks'] is not None and any(pk['sigma'] >= 1.15 for pk in r['peaks'])
            out.append(f"  Note: dividing by the light one particle scatters amplifies the small-size side and the width of the intensity "
                       f"distribution, so {w} values are far less certain than the intensity ones"
                       + ("; here a peak reaches the fit limit and dominates them" if wide or edge else "") + ".")
        return out

    # ── Parameters table ──────────────────────────────────────────────────────

    def _update_params_table(self):
        self.tree.delete(*self.tree.get_children())
        sel = self.selected[:5]
        if not sel:
            return

        # Update the column headings
        for i, path in enumerate(sel):
            n = self.samples[path]['params'].get('Sample_Name','')
            lbl = n
            self.tree.heading(f"Sample {i+1}", text=lbl)

        ROWS = [
            ("── Measurement ──", None),
            ("Date / Time", lambda p: f"{p.get('Date','')} {p.get('Time','')}"),
            ("Angle (°)", lambda p: p.get('Angle_deg')),
            ("Wavelength (nm)", lambda p: p.get('Lambda_nm')),
            ("Temperature (°C)", lambda p: p.get('Temp_C')),
            ("Viscosity (mPas)", lambda p: p.get('Visc_mPas')),
            ("Solvent n", lambda p: p.get('n_solvent')),
            ("Solvent", lambda p: p.get('Solvent')),
            ("── Results ──", None),
            ("Peak 1 (nm)", lambda p: p.get('Peak1_nm')),
            ("Mean 1 (nm)", lambda p: p.get('Mean1_nm')),
            ("SD 1 (nm)", lambda p: p.get('SD1_nm')),
            ("Area 1 (%)", lambda p: p.get('Area1_pct')),
            ("Peak 2 (nm)", lambda p: p.get('Peak2_nm')),
            ("Mean 2 (nm)", lambda p: p.get('Mean2_nm')),
            ("D10 (nm)", lambda p: p.get('D10_nm')),
            ("D50 (nm)", lambda p: p.get('D50_nm')),
            ("D90 (nm)", lambda p: p.get('D90_nm')),
            ("Mode (nm)", lambda p: p.get('Mode_nm')),
            ("Median (nm)", lambda p: p.get('Median_nm')),
            ("Z-Average (nm)", lambda p: p.get('ZAvg_nm')),
            ("PDI", lambda p: p.get('PDI')),
            ("Span", lambda p: p.get('Span')),
            ("Arithmetic mean (nm)", lambda p: p.get('AriMean_nm')),
            ("Geometric mean (nm)", lambda p: p.get('GeoMean_nm')),
        ]

        for row_i, (label, fn) in enumerate(ROWS):
            tag = "odd" if row_i % 2 else "even"
            if fn is None:
                vals = [label] + [""] * 5
                self.tree.insert("", tk.END, values=vals, tags=("section",))
                self.tree.tag_configure("section",
                                        background="#EFF6FF",
                                        font=("Helvetica", 9, "bold"))
            else:
                cells = [label]
                for path in sel:
                    v = fn(self.samples[path]['params'])
                    cells.append(str(v) if v is not None else "—")
                while len(cells) < 6:
                    cells.append("")
                self.tree.insert("", tk.END, values=cells, tags=(tag,))


    # ── Export ────────────────────────────────────────────────────────────────

    @staticmethod
    def _write_xlsx(dest_path, d):
        wb = openpyxl.Workbook()

        hdr_font  = Font(name='Arial', bold=True, size=10)
        hdr_fill  = PatternFill('solid', start_color='1E3A5F')
        hdr_font_w = Font(name='Arial', bold=True, size=10, color='FFFFFF')
        sec_fill  = PatternFill('solid', start_color='DBEAFE')
        sec_font  = Font(name='Arial', bold=True, size=10)
        cell_font = Font(name='Arial', size=10)
        thin = Side(style='thin', color='CBD5E1')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center = Alignment(horizontal='center')

        def style_header(cell, text):
            cell.value = text
            cell.font  = hdr_font_w
            cell.fill  = hdr_fill
            cell.alignment = center
            cell.border = border

        # ── ACF sheet ─────────────────────────────────────────────────────────
        ws_acf = wb.active
        ws_acf.title = 'ACF'
        headers = ['tau (µs)', 'g2-1 exp', 'g2-1 fit', 'Residual']
        for ci, h in enumerate(headers, 1):
            style_header(ws_acf.cell(1, ci), h)
        ws_acf.column_dimensions['A'].width = 14
        for col in 'BCD':
            ws_acf.column_dimensions[col].width = 13

        tau   = d['tau']
        acf   = d['acf']
        fit   = d['fit']
        resid = d['resid']
        off   = d.get('fit_offset', 6)
        n_fit, n_res = len(fit), len(resid)

        for ri, (t, g) in enumerate(zip(tau, acf), 2):
            if t <= 0:
                continue
            fi = ri - 2 - off  # index in the fit/resid vector
            ws_acf.cell(ri, 1, round(float(t), 4)).font = cell_font
            ws_acf.cell(ri, 2, round(float(g), 6)).font = cell_font
            if 0 <= fi < n_fit:
                ws_acf.cell(ri, 3, round(float(fit[fi]), 6)).font = cell_font
            if 0 <= fi < n_res:
                ws_acf.cell(ri, 4, round(float(resid[fi]), 6)).font = cell_font

        # ── Intensity distribution sheet ──────────────────────────────────────
        ws_sz = wb.create_sheet('Intensity distribution')
        for ci, h in enumerate(['Diameter (nm)', 'Intensity (%)', 'Cumulative (%)'], 1):
            style_header(ws_sz.cell(1, ci), h)
        for col in ('A','B','C'):
            ws_sz.column_dimensions[col].width = 16

        for ri, (s, iv, cv) in enumerate(zip(d['sz'], d['intens'], d['cumul']), 2):
            ws_sz.cell(ri, 1, round(float(s),  4)).font = cell_font
            ws_sz.cell(ri, 2, round(float(iv), 4)).font = cell_font
            ws_sz.cell(ri, 3, round(float(cv), 4)).font = cell_font

        # ── Number distribution sheet (Rayleigh: N ∝ I/d⁶) ────────────────────
        num_freq, num_cumul = DLSManager._to_number_dist(d['sz'], d['intens'])
        ws_nb = wb.create_sheet('Number distribution')
        for ci, h in enumerate(['Diameter (nm)', 'Number (%)', 'Cumulative (%)'], 1):
            style_header(ws_nb.cell(1, ci), h)
        for col in ('A','B','C'):
            ws_nb.column_dimensions[col].width = 16

        for ri, (s, nv, cv) in enumerate(zip(d['sz'], num_freq, num_cumul), 2):
            ws_nb.cell(ri, 1, round(float(s),  4)).font = cell_font
            ws_nb.cell(ri, 2, round(float(nv), 4)).font = cell_font
            ws_nb.cell(ri, 3, round(float(cv), 4)).font = cell_font

        # ── Parameters sheet ──────────────────────────────────────────────────
        ws_p = wb.create_sheet('Parameters')
        style_header(ws_p.cell(1, 1), 'Parameter')
        style_header(ws_p.cell(1, 2), 'Value')
        ws_p.column_dimensions['A'].width = 26
        ws_p.column_dimensions['B'].width = 22

        SECTIONS = {
            'Sample_Name': ('── Sample ──', True),
            'Angle_deg':   ('── Acquisition ──', True),
            'Peak1_nm':    ('── Results ──', True),
        }
        LABELS = {
            'Sample_Name': 'Sample name',
            'Date': 'Date',  'Time': 'Time',
            'Solvent': 'Solvent', 'Software': 'Software',
            'Angle_deg': 'Angle (°)', 'Lambda_nm': 'Wavelength (nm)',
            'Temp_C': 'Temperature (°C)', 'Visc_mPas': 'Viscosity (mPa·s)',
            'n_solvent': 'Solvent n',
            'Peak1_nm': 'Peak 1 (nm)', 'Mean1_nm': 'Mean 1 (nm)',
            'SD1_nm': 'SD 1 (nm)', 'Area1_pct': 'Area 1 (%)',
            'Peak2_nm': 'Peak 2 (nm)', 'Mean2_nm': 'Mean 2 (nm)',
            'SD2_nm': 'SD 2 (nm)',
            'D10_nm': 'D10 (nm)', 'D50_nm': 'D50 (nm)', 'D90_nm': 'D90 (nm)',
            'Mode_nm': 'Mode (nm)', 'Median_nm': 'Median (nm)',
            'ZAvg_nm': 'Z-Average (nm)', 'PDI': 'PDI',
            'Span': 'Span', 'GeoMean_nm': 'Geometric mean (nm)',
            'AriMean_nm': 'Arithmetic mean (nm)',
        }

        ri = 2
        for key, val in d['params'].items():
            if key in SECTIONS:
                sec_label, _ = SECTIONS[key]
                c1, c2 = ws_p.cell(ri, 1), ws_p.cell(ri, 2)
                c1.value = sec_label
                c1.font = sec_font; c1.fill = sec_fill; c1.border = border
                c2.fill = sec_fill; c2.border = border
                ri += 1
            label = LABELS.get(key, key)
            c1 = ws_p.cell(ri, 1, label)
            c2 = ws_p.cell(ri, 2, val if val is not None else '—')
            c1.font = cell_font; c1.border = border
            c2.font = cell_font; c2.border = border
            ri += 1

        wb.save(dest_path)

    def _export_xlsx(self):
        if not self.selected:
            messagebox.showinfo("Export", "No sample selected.")
            return
        if not XLSX_AVAILABLE:
            messagebox.showerror("Error", "The openpyxl library is not available.\nInstall it with: pip install openpyxl")
            return

        if len(self.selected) == 1:
            d = self.samples[self.selected[0]]
            dest = filedialog.asksaveasfilename(
                title="Save as...",
                initialfile=f'{d["name"]}.xlsx',
                defaultextension='.xlsx',
                filetypes=[("Excel", "*.xlsx")])
            if dest:
                self._write_xlsx(dest, d)
                messagebox.showinfo("Export completed", f"Saved:\n{dest}")
        else:
            folder = filedialog.askdirectory(title="Destination folder")
            if not folder:
                return
            saved = []
            for path in self.selected:
                d = self.samples[path]
                out = os.path.join(folder, f'{d["name"]}.xlsx')
                self._write_xlsx(out, d)
                saved.append(os.path.basename(out))
            messagebox.showinfo("Export completed",
                                f"Saved {len(saved)} files in:\n{folder}\n\n" +
                                "\n".join(saved))

    # ── Figure (PlotStyleKit) ─────────────────────────────────────────────────

    def _current_figure(self):
        """The figure of the active plot tab, or None on the Parameters tab or with nothing selected."""
        if not self.selected:
            return None
        tab = self.notebook.select()
        if tab == str(self.tab_acf):
            return self.fig_acf
        if tab == str(self.tab_size):
            return self.fig_size
        if tab == str(self.tab_dls):
            return self.fig_dls
        return None

    @staticmethod
    def _preset_for(fig):
        """'single' (4:3) for one panel, 'double' (16:9) for two."""
        return 'single' if len(fig.axes) == 1 else 'double'

    def _copy_figure(self, fig):
        """Copy of `fig` for saving/editor, restyled Origin by PlotStyleKit when it is in use ('single' 4:3 for one
        panel, 'double' 16:9 for two), else as it is."""
        if psk is not None:
            return psk.publication_copy(fig, self._preset_for(fig), unpin=False)
        return pickle.loads(pickle.dumps(fig))

    def _figure_or_warn(self, title):
        fig = self._current_figure()
        if fig is None:
            messagebox.showwarning(title, "Select samples and open the Autocorrelogram, Distribution or DLS analysis tab.")
        return fig

    def save_figure_image(self):
        fig = self._figure_or_warn("Save Figure Image")
        if fig is None:
            return
        if psk is not None:
            psk.save_image_dialog(fig, self.root, preset=self._preset_for(fig), unpin=False)
            return
        path = filedialog.asksaveasfilename(
            title="Save Figure Image", defaultextension=".png",
            filetypes=[("PNG", "*.png"), ("PDF", "*.pdf"), ("SVG", "*.svg"), ("All Files", "*.*")])
        if not path:
            return
        try:
            self._copy_figure(fig).savefig(path, dpi=300)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Save Figure Image", f"Save failed:\n{e}")
            return
        messagebox.showinfo("Save Figure Image", f"Figure saved (plain look: PlotStyleKit is not in use):\n{path}")

    def save_figure_pickle(self):
        """Save the figure as a matplotlib pickle: reopenable with plot_editor.pyw as a live Figure object."""
        fig = self._figure_or_warn("Save Figure")
        if fig is None:
            return
        if psk is not None:
            psk.save_pickle_dialog(fig, self.root, preset=self._preset_for(fig), unpin=False)
            return
        path = filedialog.asksaveasfilename(
            title="Save Figure", defaultextension=".fig.pickle",
            filetypes=[("Matplotlib Figure (pickle)", "*.pickle *.pkl"), ("All Files", "*.*")])
        if not path:
            return
        try:
            with open(path, 'wb') as f:
                pickle.dump(self._copy_figure(fig), f)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Save Figure", f"Save failed:\n{e}")
            return
        messagebox.showinfo("Save Figure", f"Figure saved (not restyled: PlotStyleKit is not in use):\n{path}")

    def open_figure_editor(self):
        """Open the PlotStyleKit Plot Editor on a copy of the current figure."""
        fig = self._figure_or_warn("Edit Figure")
        if fig is None:
            return
        if not self._need_plotstylekit("Edit Figure"):
            return
        psk.edit_dialog(fig, self.root, title="current figure", preset=self._preset_for(fig), unpin=False)

    # ---- PlotStyleKit: the Origin look and the figure editor (a fallback without it for each) ----
    def _adopt_plotstylekit(self, mod):
        """Start a PlotStyleKit that was just located (at start-up or from File > PlotStyleKit...)."""
        global psk, PSK_PROBLEM
        try:
            mod.setup()
        except Exception as e:
            messagebox.showwarning("PlotStyleKit", f"PlotStyleKit failed to start:\n{e}", parent=self.root)
            return False
        psk, PSK_PROBLEM = mod, None
        self._refresh_all()
        messagebox.showinfo("PlotStyleKit", f"PlotStyleKit {mod.VERSION} is in use:\n{os.path.dirname(mod.__file__)}",
                            parent=self.root)
        return True

    def _offer_plotstylekit(self, problem):
        """The warning that asks for the folder of PlotStyleKit, with a Cancel for going without it."""
        mod = _ask_plotstylekit_folder(self.root, problem)
        if mod is not None:
            self._adopt_plotstylekit(mod)

    def plotstylekit_dialog(self):
        """File > PlotStyleKit...: locate it (when it is not in use) or pick another copy (when it is)."""
        if psk is None:
            _psk_save(declined=False)
            self._offer_plotstylekit(PSK_PROBLEM or "PlotStyleKit is not in use")
            return
        folder = filedialog.askdirectory(parent=self.root, title="Folder of PlotStyleKit (now in use: "
                                         f"{os.path.dirname(psk.__file__)})")
        if not folder:
            return
        mod, why = _psk_try(folder, PSK_MIN_API)
        if mod is None:
            messagebox.showwarning("PlotStyleKit", why + ".", parent=self.root)
            return
        _psk_save(path=folder, declined=False)
        self._adopt_plotstylekit(mod)

    def _need_plotstylekit(self, what):
        """For what only PlotStyleKit does (the figure editor): offer to find it. True if it is in use afterwards."""
        if psk is None and messagebox.askyesno(
                what, f"{what} belongs to PlotStyleKit, which is not in use.\n\nLocate PlotStyleKit now?",
                parent=self.root):
            self.plotstylekit_dialog()
        return psk is not None

    # ── Utility ───────────────────────────────────────────────────────────────

    def _short_label(self, name):
        if len(name) > 35:
            return name[:32] + '…'
        return name


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    import sys

    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        print("Note: tkinterdnd2 not available, drag & drop disabled.")

    app = DLSManager(root)

    if not OLEFILE_AVAILABLE:
        messagebox.showwarning("Missing dependency",
                               "The olefile library is not available: .nsz files cannot be opened.\n"
                               "Install it with: pip install olefile")

    # If files are passed on the command line, load them
    if len(sys.argv) > 1:
        files = [a for a in sys.argv[1:] if a.lower().endswith('.nsz')]
        if files:
            app._load_files(files)

    root.mainloop()


if __name__ == '__main__':
    main()
