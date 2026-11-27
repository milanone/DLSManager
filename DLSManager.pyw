"""
DLS Manager — Horiba SZ-100 DLS: .nsz viewer, correlogram analysis and size distributions
Supports drag & drop, multiple-sample comparison, Excel export.
The plots follow the shared Origin style of PlotStyleKit when the sibling repo is present.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, struct, re, csv, io, pickle, traceback
from html.parser import HTMLParser
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

def _csv_number(text):
    """The first number in a cell of the software's CSV ('153.2', '1.590 - 0.000i'), or None for '---' and empty cells."""
    m = re.match(r"\s*(-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)", text or "")
    return float(m.group(1)) if m else None


def load_csv(path):
    """Load a CSV export of the SZ-100 software (a measurement exported, or a recalculation: it is the only file a recalculation
    writes, no new .nsz) as a sample dict shaped like load_nsz(): the size table (its diameters are the bins' upper edges, as in
    the .nsz), the results, the correlation g1 with its fit and the form of distribution named in the header. The raw
    correlogram (g2-1 from the first channel) is not in the CSV, so the Recomputed source is unavailable for these samples
    (dls_record() returns None for 'from_csv'); the correlation is stored as g2-1 = g1^2 like the .nsz one, from the first delay the CSV lists."""
    with open(path, newline="", encoding="latin-1") as f:
        rows = [r for r in csv.reader(f)]
    if not rows or not rows[0] or rows[0][0].strip() != "Display Setting":
        raise ValueError("not an export of the SZ-100 software")
    items = {r[0].strip(): r[1].strip() for r in rows if len(r) >= 2 and r[0].strip() and not r[0].strip()[0].isdigit()}

    def table(header):
        start = next((i for i, r in enumerate(rows) if r and ",".join(r).startswith(header)), None)
        if start is None:
            return np.zeros((0, 0))
        out = []
        for r in rows[start + 1:]:
            if not r or not r[0].strip():
                break
            out.append([float(x) for x in r])
        return np.array(out)

    size, corr = table("Diameter (nm)"), table("Delay Time(usec),Correlation")
    if size.shape[0] < 10 or corr.shape[0] < 8:
        raise ValueError("the size table or the correlation is missing")
    g1, fit_g1 = corr[:, 1], corr[:, 2]
    used = int(np.argmax(fit_g1 == 0)) if np.any(fit_g1 == 0) else len(fit_g1)       # the fit is 0 outside the channels used
    num = lambda k: _csv_number(items.get(k))
    date = re.sub(r"^\D+", "", items.get("Date", ""))            # without the weekday ('<weekday> 19 novembre 2026 18:37:34')
    parts = date.rsplit(" ", 1)
    form = items.get("Distribution Form"), items.get("Distribution Form(Dispersity)")
    peak_mode = num("Peak1 Mode")
    params = {
        'Sample_Name': items.get("Sample Name", ""),
        'Date': parts[0] if len(parts) == 2 else date, 'Time': parts[1] if len(parts) == 2 else '',
        'Lambda_nm': 532.0, 'Angle_deg': num("Scattering Angle"),   # the SZ-100's laser; the CSV does not state it
        'Temp_C': num("Temperature of the Holder"),
        'Visc_mPas': num("Dispersion Medium Viscosity"), 'n_solvent': num("Dispersion Medium RI"),
        'Solvent': items.get("Dispersion Medium Name", ""), 'n_particle': num("Particle Refractive Index"),
        'Peak1_nm': peak_mode, 'Mode_nm': peak_mode, 'Mean1_nm': num("Peak1 Size (Mean)"), 'AriMean_nm': num("Peak1 Size (Mean)"), 'SD1_nm': num("Peak1 S. D."),
        'Area1_pct': num("Peak1 S.P.Area Ratio"), 'Mean2_nm': num("Peak2 Size (Mean)"), 'Peak2_nm': num("Peak2 Mode"),
        'ZAvg_nm': num("Z-Average"), 'PDI': num("PI"), 'CumSD_nm': num("Cumulants S. D."),
        'FormText': ", ".join(x for x in (form[1], form[0]) if x) or None,         # 'Monodisperse, Narrow', as the header says
        'Representation': items.get("Representation of Result"),
    }
    return {
        'path': path, 'name': os.path.splitext(os.path.basename(path))[0],
        'tau': corr[:, 0], 'acf': g1 ** 2, 'fit': fit_g1[:used] ** 2, 'resid': [], 'fit_offset': 0, 'tau_aligned': True,
        'sz': size[:, 0], 'intens': size[:, 1], 'cumul': size[:, 2], 'params': params, 'from_csv': True,
    }


AVERAGED_PARAMS = ('Temp_C', 'Visc_mPas', 'Peak1_nm', 'Mean1_nm', 'SD1_nm', 'Area1_pct', 'Peak2_nm', 'Mean2_nm', 'SD2_nm',
                   'D10_nm', 'D50_nm', 'D90_nm', 'Mode_nm', 'Median_nm', 'ZAvg_nm', 'PDI', 'CumSD_nm', 'Span', 'GeoMean_nm',
                   'AriMean_nm')
# conditions that must be the same for two measurements to be averaged: (key, name in the message, tolerance)
SAME_FOR_AVERAGE = (('Angle_deg', "scattering angle", 0.5), ('Lambda_nm', "wavelength", 0.5), ('n_solvent', "solvent index", 0.002),
                    ('n_particle', "particle index", 0.002), ('Temp_C', "temperature", 0.5), ('Visc_mPas', "viscosity", 0.01),
                    ('Resolution', "form of distribution", 1e-3), ('Dispersity', "form of distribution", 1e-3),
                    ('FormText', "form of distribution", None), ('Range_from', "first channel of the distribution", 0.5))


def dls_average(samples):
    """The mean of two or more measurements (sample dicts of load_nsz() / load_csv()) as a new sample: correlogram channel by
    channel (so the Recomputed source inverts the averaged correlogram), the instrument's size table bin by bin, its fit over
    the channels common to all, and each numeric result as the mean of the measurements' values, with their standard deviation
    in 'spread'. ValueError, with the reason, if they cannot be averaged: other conditions or form of distribution, other
    delays or size bins, a mix of .nsz and CSV, or an average among them (it would weigh its measurements twice or unequally)."""
    if len(samples) < 2:
        raise ValueError("At least two measurements are needed.")
    if any(d.get('sources') for d in samples):
        raise ValueError("An average cannot be averaged again: the original measurements are to be selected.")
    first = samples[0]
    if len({bool(d.get('from_csv')) for d in samples}) > 1:
        raise ValueError(".nsz files and CSV exports cannot be averaged together (their correlograms start at other delays).")
    if not all(d.get('tau_aligned') for d in samples):
        raise ValueError("A file has delays that are not matched to its correlogram.")
    for key, what in (('tau', "correlator delays"), ('sz', "size bins")):
        for d in samples[1:]:
            if len(d[key]) != len(first[key]) or not np.allclose(d[key], first[key], rtol=1e-3):
                raise ValueError(f"The {what} of '{d['name']}' differ from those of '{first['name']}'.")
    for key, what, tol in SAME_FOR_AVERAGE:
        a = first['params'].get(key)
        for d in samples[1:]:
            b = d['params'].get(key)
            same = a == b if tol is None or a is None or b is None else abs(a - b) <= tol
            if not same:
                raise ValueError(f"The {what} of '{d['name']}' ({b}) differs from that of '{first['name']}' ({a}).")

    n = len(samples)
    params = dict(first['params'])
    spread = {}
    for key in AVERAGED_PARAMS:
        values = [d['params'].get(key) for d in samples]
        if any(v is None for v in values):
            params[key] = None
            continue
        params[key] = round(float(np.mean(values)), 4)
        spread[key] = float(np.std(values, ddof=1))
    ends = [d['params'].get('Range_to') for d in samples]
    params['Range_to'] = None if any(e is None for e in ends) else min(ends)      # the channels every measurement used
    when = sorted((d['params'].get('Date', ''), d['params'].get('Time', '')) for d in samples)
    params['Date'], params['Time'] = when[0][0], f"{when[0][1]} to {when[-1][1]}"
    params['Sample_Name'] = f"{first['params'].get('Sample_Name') or first['name']} (average of {n})"

    offsets = {d.get('fit_offset', 6) for d in samples}
    def mean_over_common(key):
        length = min(len(d[key]) for d in samples)
        if len(offsets) > 1 or length < 3:
            return np.array([])
        return np.mean([np.asarray(d[key][:length], float) for d in samples], axis=0)

    parts = [re.match(r"(.*?)(?:_(\d+))?$", d['name']).groups() for d in samples]
    numbers = [number for _, number in parts]
    name = f"{parts[0][0]}_avg_{'+'.join(numbers)}" if all(numbers) else f"{parts[0][0]}_avg{n}"
    return {
        'name': name, 'path': os.path.join(os.path.dirname(first['path']), name + ".avg"),
        'tau': first['tau'].copy(), 'acf': np.mean([d['acf'] for d in samples], axis=0),
        'fit': mean_over_common('fit'), 'resid': mean_over_common('resid'), 'fit_offset': first.get('fit_offset', 6),
        'tau_aligned': True, 'from_csv': bool(first.get('from_csv')),
        'sz': first['sz'].copy(), 'intens': np.mean([d['intens'] for d in samples], axis=0),
        'cumul': np.mean([d['cumul'] for d in samples], axis=0),
        'params': params, 'spread': spread, 'sources': [d['path'] for d in samples],
    }


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
        'Resolution': pv(p1, 'Dist_Resolution'),   # 0.3 in every file seen (Monodisperse Standard)
        'Dispersity': pv(p34, 'MeasDispersityMode'),   # 0 = Monodisperse (the only value seen)
        'CumSD_nm': pv(p1, 'Cum_fStd',   2.0),  # = Z-average x sqrt(PI) exactly (28 files); a radius, x2
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
    'shell': "Mie shell: a vesicle as a shell around a core filled with the medium. Lamellae: number of bilayers (1 = "
             "unilamellar). Bilayer: thickness of one bilayer, about 4.5 nm. Repeat distance: bilayer plus the water "
             "between two lamellae, about 7 nm. The lamellae are treated as one shell of n x repeat - (repeat - bilayer) "
             "whose index excess is the lipid's (0.12, an order-of-magnitude value) averaged over the water gaps. The "
             "diameter is the outer one. The light one vesicle scatters grows about as d^4 at small sizes instead of "
             "d^6. Only used by the 'Mie shell' conversion.",
    'width': "Free (regularised) distribution only, all weightings. As computed: the distribution as the inversion gives it. "
             "Broad, Standard and Narrow are the forms of the instrument software (Edit > Recalculate, Form of Distribution): "
             "the instrument shows a peak of a FIXED width whatever the sample (sigma of ln d 0.43, 0.24 and 0.07; its "
             "resolution -0.5, 0.3 and 0.6), and its own data are always shown in the Standard form. Here the computed "
             "distribution is given that width around its centre, in intensity and before the conversion to volume and "
             "number. Manual: a resolution from -1 to 1 (the instrument's own entry), mapped to a width through the four "
             "points measured on its recalculated tables. The regularised inversion widens a narrow sample, and volume "
             "and number depend strongly on the width, so a narrow form gives numbers closer to a narrow sample (a "
             "latex of 10% CV: number mean 130 nm); for a truly broad sample it narrows what is really wide, and the "
             "curve no longer follows the correlogram fit drawn beside it. The instrument's Polydisperse form allows "
             "several peaks: the two-peak model is the nearest equivalent.",
    'resolution': "Manual form only: the instrument's resolution, from -1 to 1 (larger = narrower peak). Presets of its "
                  "manual: Broad -0.5, Standard 0.3, Narrow 0.6. The width it gives is read from the points measured on "
                  "the instrument's recalculated size tables (-0.5: 0.43, 0.1: 0.30, 0.3: 0.24, 0.6: 0.07, sigma of ln d).",
    'aperture': "Half-width of the angular range over which the detector collects light, degrees (Mie and Mie shell). "
                "The light one particle scatters is averaged over the scattering angle +- this value, which washes out "
                "the narrow interference minima of a single angle; without it, a small tail of the distribution sitting "
                "at such a minimum becomes a spurious peak in volume and number. The value for the SZ-100 is not known: "
                "the default is 0 (the single angle), at most 45. The minimum is only removed by an aperture of about 30 "
                "degrees or more, wider than a real detector: it is a smoothing, not a physical correction.",
    'pn': "Refractive index of the particle for the 'Mie' conversion (a solid sphere). Empty: the value stored in the "
          "file, set in the instrument software (1.6 by default). Not used by 'Mie shell' (which has its own "
          "parameters) or 'Rayleigh'.",
    'source': "Instrument: the size distribution stored in the file by the instrument software (with its cumulative "
              "curve). Recomputed: a new analysis of the stored correlogram, with the model chosen below (cumulants, "
              "free distribution or one or two log-normal peaks), shown next to the instrument's values.",
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
    (diameters nm, p normalised to sum 1, fitted g1 at all delays, with the intercept free). Needs scipy (nnls)."""
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
    # the sum of the weights is the intercept of g1 at zero delay: it is fitted, not forced to 1 (the stored ACF starts
    # near 0.99), so the fit below is p unnormalised; only the distribution that is returned is normalised
    return d, p / p.sum(), np.exp(-np.outer(tau, gam)) @ p


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
SIZE_SOURCES = ("Instrument", "Recomputed")
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


DLS_APERTURE = 0.0       # default half-width (degrees) of the detector's angular acceptance: 0 = the single angle (the SZ-100's value is unknown)


def dls_per_particle(d, conversion, rec, wavelength, n_particle, shell=(4.5, 0.12), aperture=0.0):
    """Intensity scattered by one particle of each diameter `d`: d^6 for 'Rayleigh', Mie theory for 'Mie' (needs the
    scattering angle and solvent index of `rec`, the laser wavelength and the refractive index of the particle).
    `aperture` (degrees, half-width) averages the Mie intensity over the angles angle +- aperture (9 angles): a detector
    collects light over an angular range, which washes out the sharp interference minima of a single angle (at 90 degrees
    a latex of 350 nm scatters 5 times less than one of 250 nm, and dividing by that makes a spurious large peak out of the
    small tail of a distribution). 0 is the single angle. Rayleigh is not affected."""
    d = np.asarray(d, float)
    if conversion == "Rayleigh":
        return d ** 6
    angles = [rec['angle']] if aperture <= 0 else list(np.clip(rec['angle'] + np.linspace(-aperture, aperture, 9), 0.0, 180.0))
    if conversion == "Mie shell":
        return np.mean([mie_shell_intensity(d, wavelength, rec['n'], shell[1], shell[0], a) for a in angles], axis=0)
    return np.mean([mie_intensity(d, wavelength, rec['n'], n_particle, a) for a in angles], axis=0)


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


# The instrument's own forms (Edit > Recalculate, 'Form of Distribution'): a 'resolution' from -1 to 1 fixes the WIDTH of the peaks it
# shows, whatever the data (manual: Monodisperse Broad -0.5, Standard 0.3, Narrow 0.6; Polydisperse -0.3, 0.1, 0.3). Measured on its
# recalculated size tables (sigma of ln d): Monodisperse Broad 0.42-0.45, Standard 0.23-0.25, Narrow 0.07; Polydisperse Standard 0.30 on a
# latex. Piecewise-linear map through these four points (resolution, sigma), extended linearly beyond them; an interpolation of few
# points, an assumption outside them.
DLS_RESOLUTION_SIGMA = ((-0.5, 0.43), (0.1, 0.30), (0.3, 0.24), (0.6, 0.07))
DLS_FORM_RESOLUTION = {"Broad": -0.5, "Standard": 0.3, "Narrow": 0.6}       # the Monodisperse presets of the manual
DLS_POLY_RESOLUTION = {"Broad": -0.3, "Standard": 0.1, "Narrow": 0.3}      # the Polydisperse presets of the manual
DLS_WIDTHS = ("As computed", "Broad", "Standard", "Narrow", "Manual")


def dls_form_name(dispersity, resolution):
    """'Monodisperse, Standard' for the form of distribution a file was calculated with, from its dispersity code (0 = Monodisperse,
    the only one seen; anything else is reported as a code, not guessed) and its resolution, matched to the manual's presets
    (Monodisperse -0.5 / 0.3 / 0.6, Polydisperse -0.3 / 0.1 / 0.3 for Broad / Standard / Narrow); another resolution is 'Manual'.
    None if the file has no resolution."""
    if resolution is None:
        return None
    if not dispersity:
        kind, presets = "Monodisperse", DLS_FORM_RESOLUTION
    else:
        kind, presets = "Dispersity code %d" % int(dispersity), {}
    width = next((n for n, r in presets.items() if abs(r - resolution) < 1e-3), "Manual")
    return f"{kind}, {width}"


def dls_resolution_sigma(resolution):
    """Width (sigma of ln d) of the peak the instrument shows for a resolution setting from -1 to 1 (see DLS_RESOLUTION_SIGMA)."""
    xs, ys = zip(*DLS_RESOLUTION_SIGMA)
    if resolution <= xs[0]:
        s = ys[0] + (ys[1] - ys[0]) / (xs[1] - xs[0]) * (resolution - xs[0])
    elif resolution >= xs[-1]:
        s = ys[-1] + (ys[-1] - ys[-2]) / (xs[-1] - xs[-2]) * (resolution - xs[-1])
    else:
        s = float(np.interp(resolution, xs, ys))
    return float(np.clip(s, 0.03, 1.0))


def _rescale_width(d, p, factor):
    """`p` with its width in ln d multiplied by `factor` (any positive value: below 1 narrows, above 1 widens) around its
    mean ln d, normalised to sum 1."""
    lnd = np.log(np.asarray(d, float))
    m = np.sum(p * lnd) / np.sum(p)
    q = np.interp(m + (lnd - m) / factor, lnd, p, left=0.0, right=0.0)
    return q / q.sum()


def dls_set_width(d, p, sigma):
    """`p` (weights on the log grid `d`) given a width (sigma of ln d) of exactly `sigma` around its mean ln d, narrower or wider
    than computed, normalised to sum 1. It imitates the instrument's forms, whose distribution has a fixed width (about 0.24 in
    its Standard form) whatever the polydispersity of the sample: an assumption of the instrument, not information about the
    sample. The regularised inversion widens a narrow sample (sigma 0.31 for a true 0.10 on a synthetic latex), and number and
    volume depend on the width through about exp(-k sigma^2)."""
    p = np.asarray(p, float)
    lnd = np.log(np.asarray(d, float))
    m = np.sum(p * lnd) / np.sum(p)
    cur = float(np.sqrt(np.sum(p * (lnd - m) ** 2) / np.sum(p)))
    return _rescale_width(d, p, sigma / cur) if cur > 0 else p / p.sum()


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


def rec_wl(d):
    """Laser wavelength (nm) of a sample: the one stored in the file, else 532 (SZ-100)."""
    return d['params'].get('Lambda_nm') or 532.0


def dls_record(d):
    """The input of dls_analyse() for a sample dict of load_nsz(): delays in seconds and the measurement conditions
    (angle, solvent index, viscosity, temperature) read from the file. None if the file lacks any of them or the delays
    are not matched to the channels."""
    p = d['params']
    keys = ('Angle_deg', 'n_solvent', 'Visc_mPas', 'Temp_C')
    if not d.get('tau_aligned') or d.get('from_csv') or any(p.get(k) is None for k in keys) or len(d['tau']) < 8:
        return None
    rec = {'tau': d['tau'] * 1e-6, 'acf': d['acf'], 'angle': p['Angle_deg'], 'n': p['n_solvent'],
           'visc': p['Visc_mPas'], 'T': p['Temp_C']}
    if p.get('Range_from'):
        rec['range'] = (p['Range_from'] - 1, p['Range_to'])      # 0-based [first, end) channel indices
    return rec


class _GuideParser(HTMLParser):
    """Reads the guide (docs/sz100_guide.html) into blocks for GuideWindow: ('h1'|'h2'|'h3'|'p'|'li'|'note', runs, id) and
    ('table', rows, None), where `runs` is a list of (text, style) with style '', 'b', 'i' or 'code' and a row is a list of
    (is_header, runs). The navigation, the style sheet and the title are skipped (the window has its own contents list)."""
    SKIP = ("nav", "style", "title", "script")
    INLINE = {"b": "b", "strong": "b", "i": "i", "em": "i", "code": "code"}

    def __init__(self):
        super().__init__()
        self.blocks, self.cur, self.kind, self.bid, self.skip = [], None, None, None, 0
        self.fmt, self.rows, self.row, self.cell_header = [], None, None, False

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in self.SKIP:
            self.skip += 1
        elif self.skip:
            return
        elif tag in ("h1", "h2", "h3", "p", "li"):
            self._begin(tag, a.get("id"))
        elif tag == "div" and "note" in (a.get("class") or ""):
            self._begin("note", None)
        elif tag == "table":
            self.rows = []
        elif tag == "tr" and self.rows is not None:
            self.row = []
        elif tag in ("td", "th") and self.row is not None:
            self.cell_header = tag == "th"
            self._begin("cell", None)
        elif tag in self.INLINE:
            self.fmt.append(self.INLINE[tag])
        elif tag == "br" and self.cur is not None:
            self.cur.append(("\n", ""))

    def handle_endtag(self, tag):
        if tag in self.SKIP:
            self.skip = max(0, self.skip - 1)
        elif self.skip:
            return
        elif tag in ("h1", "h2", "h3", "p", "li") or (tag == "div" and self.kind == "note"):
            self._end()
        elif tag in ("td", "th") and self.kind == "cell":
            self.row.append((self.cell_header, self._trim(self.cur)))
            self.cur, self.kind = None, None
        elif tag == "tr" and self.row is not None:
            self.rows.append(self.row)
            self.row = None
        elif tag == "table" and self.rows is not None:
            self.blocks.append(("table", self.rows, None))
            self.rows = None
        elif tag in self.INLINE and self.INLINE[tag] in self.fmt:
            self.fmt.remove(self.INLINE[tag])

    def handle_data(self, data):
        if self.skip or self.cur is None:
            return
        text = re.sub(r"\s+", " ", data)
        if text:
            self.cur.append((text, self.fmt[-1] if self.fmt else ""))

    def _begin(self, kind, bid):
        self.cur, self.kind, self.bid = [], kind, bid

    def _end(self):
        if self.cur is not None:
            runs = self._trim(self.cur)
            if runs:
                self.blocks.append((self.kind, runs, self.bid))
        self.cur, self.kind = None, None

    @staticmethod
    def _trim(runs):
        runs = [list(r) for r in (runs or [])]
        while runs and not runs[0][0].strip():
            runs.pop(0)
        while runs and not runs[-1][0].strip():
            runs.pop()
        if runs:
            runs[0][0], runs[-1][0] = runs[0][0].lstrip(), runs[-1][0].rstrip()
        return [tuple(r) for r in runs]


class GuideWindow(tk.Toplevel):
    """The guide to the SZ-100 settings, shown in a window of the program (not in a browser): a contents list on the left and the
    text on the right, read from docs/sz100_guide.html. Tables of short cells are aligned in a monospaced font; those with long
    cells are written one row at a time, the first cell in bold and the others under their column titles."""

    def __init__(self, master, path):
        super().__init__(master)
        self.title("SZ-100 settings and recalculation (guide)")
        self.geometry("1000x680")
        with open(path, encoding="utf-8") as f:
            parser = _GuideParser()
            parser.feed(f.read())
        self.sections = []                               # (title, mark name) of the h2 headings
        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True)
        left = ttk.Frame(paned)
        self.contents = tk.Listbox(left, width=34, activestyle="none", exportselection=False, borderwidth=0, highlightthickness=0)
        self.contents.pack(fill="both", expand=True, padx=(6, 0), pady=6)
        right = ttk.Frame(paned)
        self.text = tk.Text(right, wrap="word", padx=16, pady=10, borderwidth=0, highlightthickness=0, cursor="arrow",
                            font=("Segoe UI", 10), spacing3=4, background="#ffffff", foreground="#1a1a1a")
        sb = ttk.Scrollbar(right, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)
        paned.add(left, weight=0)
        paned.add(right, weight=1)
        self._tags()
        self._render(parser.blocks)
        self.text.configure(state="disabled")
        self.contents.bind("<<ListboxSelect>>", self._jump)
        self.text.bind("<Button-1>", lambda e: self.text.focus_set())

    def _tags(self):
        t = self.text
        t.tag_configure("h1", font=("Segoe UI", 16, "bold"), spacing3=8)
        t.tag_configure("h2", font=("Segoe UI", 13, "bold"), spacing1=18, spacing3=6, foreground="#1f5fa8")
        t.tag_configure("h3", font=("Segoe UI", 10, "bold"), spacing1=8, spacing3=2)
        t.tag_configure("p", lmargin1=0, lmargin2=0)
        t.tag_configure("li", lmargin1=14, lmargin2=28)
        t.tag_configure("note", lmargin1=12, lmargin2=12, rmargin=12, background="#eef3f8", spacing1=6, spacing3=8)
        t.tag_configure("b", font=("Segoe UI", 10, "bold"))
        t.tag_configure("i", font=("Segoe UI", 10, "italic"))
        t.tag_configure("code", font=("Consolas", 10), background="#f0f2f5")
        t.tag_configure("mono", font=("Consolas", 10), lmargin1=10, lmargin2=10, spacing3=0)
        t.tag_configure("monohead", font=("Consolas", 10, "bold"), lmargin1=10, lmargin2=10, spacing3=0, background="#eef3f8")
        t.tag_configure("rowfirst", font=("Segoe UI", 10, "bold"), lmargin1=10, lmargin2=10, spacing1=6, spacing3=0)
        t.tag_configure("rowcell", lmargin1=26, lmargin2=26, spacing3=0)
        t.tag_configure("rowtitle", font=("Segoe UI", 9, "italic"), foreground="#5b6470")
        t.tag_raise("b"); t.tag_raise("i"); t.tag_raise("code")        # an inline style wins over the block's font

    @staticmethod
    def _plain(runs):
        return "".join(r[0] for r in runs)

    def _runs(self, runs, base):
        for text, style in runs:
            self.text.insert("end", text, (base, style) if style else (base,))

    def _render(self, blocks):
        for kind, content, bid in blocks:
            if kind == "table":
                self._table(content)
                continue
            if kind == "h2":
                mark = "sec%d" % len(self.sections)
                self.text.mark_set(mark, "end-1c")
                self.text.mark_gravity(mark, "left")
                self.sections.append((self._plain(content), mark))
                self.contents.insert("end", self._plain(content))
            if kind == "li":
                self.text.insert("end", "•  ", ("li",))
            self._runs(content, kind)
            self.text.insert("end", "\n", (kind,))

    def _table(self, rows):
        cells = [[self._plain(r) for _, r in row] for row in rows]
        cols = max(len(r) for r in cells)
        widths = [max(len(r[c]) if c < len(r) else 0 for r in cells) for c in range(cols)]
        if sum(widths) + 3 * cols <= 84:                 # short cells: aligned columns
            for row, cs in zip(rows, cells):
                is_head = all(h for h, _ in row)
                line = "   ".join(cs[c].ljust(widths[c]) if c < len(cs) else " " * widths[c] for c in range(cols))
                self.text.insert("end", line.rstrip() + "\n", ("monohead" if is_head else "mono",))
            self.text.insert("end", "\n")
            return
        head = cells[0] if all(h for h, _ in rows[0]) else None
        for row in rows[1 if head else 0:]:
            self._runs(row[0][1], "rowfirst")
            self.text.insert("end", "\n", ("rowfirst",))
            for c, (_, runs) in enumerate(row[1:], 1):
                if head and c < len(head):
                    self.text.insert("end", head[c] + ": ", ("rowcell", "rowtitle"))
                self._runs(runs, "rowcell")
                self.text.insert("end", "\n", ("rowcell",))
        self.text.insert("end", "\n")

    def _jump(self, _event=None):
        sel = self.contents.curselection()
        if sel:
            self.text.yview(self.sections[sel[0]][1])


class PlotCursor:
    """A movable vertical bar over a figure, to read the values of the curves. Left click or drag in a panel puts the bar
    there (the pan/zoom tools of the toolbar take precedence while they are active); right click removes it. A box
    in the panel lists, for every labelled curve, its value at the bar (interpolated; a curve that does not reach the bar
    is skipped). The figure is cleared and redrawn by the plot methods, so they call attach() afterwards, which puts
    the bar back at the same place; remove() takes it out of the figure (before saving or editing a copy)."""

    def __init__(self, canvas):
        self.canvas, self.fig = canvas, canvas.figure
        self.x = None              # position in data coordinates
        self.panel = None          # index of the axes that holds the bar
        self.line = self.box = None
        self._drag = False
        canvas.mpl_connect('button_press_event', self._press)
        canvas.mpl_connect('motion_notify_event', self._move)
        canvas.mpl_connect('button_release_event', self._release)

    def _tool_active(self):
        mode = getattr(getattr(self.canvas, 'toolbar', None), 'mode', '')
        return bool(getattr(mode, 'value', mode))

    def remove(self):
        for art in (self.line, self.box):
            try:
                if art is not None:
                    art.remove()
            except (ValueError, NotImplementedError, AttributeError):
                pass
        self.line = self.box = None

    def attach(self):
        """Put the bar back after the figure was rebuilt; forget it if its panel or position no longer exists."""
        self.line = self.box = None
        if self.x is None or self.panel is None:
            return
        if self.panel >= len(self.fig.axes) or not self._inside(self.fig.axes[self.panel], self.x):
            self.x = self.panel = None
            return
        self._create(self.fig.axes[self.panel])

    @staticmethod
    def _inside(ax, x):
        lo, hi = sorted(ax.get_xlim())
        return lo <= x <= hi

    def _create(self, ax):
        self.line = ax.axvline(self.x, color='#DC2626', lw=1.0, alpha=0.85, zorder=10)
        self.box = ax.text(0.02, 0.98, '', transform=ax.transAxes, va='top', ha='left', fontsize=8, zorder=11,
                           bbox=dict(boxstyle='round,pad=0.3', fc='white', ec='#94A3B8', alpha=0.92))
        self._update(ax)

    def readout(self, ax):
        """Text of the box: the position and the value of each labelled curve at the bar."""
        rows = ["x = %.4g" % self.x]
        for ln in ax.lines:
            if ln is self.line:
                continue
            name = ln.get_label()
            if name.startswith('_'):
                name = getattr(ln, '_cursor_name', None)
            xd, yd = np.asarray(ln.get_xdata(), float), np.asarray(ln.get_ydata(), float)
            if not name or len(xd) < 3 or len(xd) != len(yd):
                continue
            ok = np.isfinite(xd) & np.isfinite(yd)
            xd, yd = xd[ok], yd[ok]
            if len(xd) < 3 or not xd.min() <= self.x <= xd.max():
                continue
            order = np.argsort(xd)
            rows.append("%s: %.4g" % (name[:40], np.interp(self.x, xd[order], yd[order])))
        return "\n".join(rows[:9])

    def _update(self, ax):
        self.line.set_xdata([self.x, self.x])
        self.box.set_text(self.readout(ax))

    def _press(self, event):
        if event.inaxes is None or self._tool_active() or event.xdata is None:
            return
        if event.button == 3:                      # right click: remove the bar
            self.remove()
            self.x = self.panel = None
            self.canvas.draw_idle()
            return
        if event.button != 1:
            return
        panel = self.fig.axes.index(event.inaxes)
        self.x = float(event.xdata)
        if self.line is None or panel != self.panel:
            self.remove()
            self.panel = panel
            self._create(event.inaxes)
        else:
            self._update(event.inaxes)
        self._drag = True
        self.canvas.draw_idle()

    def _move(self, event):
        if not self._drag or event.inaxes is None or event.xdata is None or self.line is None:
            return
        if self.fig.axes.index(event.inaxes) != self.panel:
            return
        self.x = float(event.xdata)
        self._update(event.inaxes)
        self.canvas.draw_idle()

    def _release(self, event):
        self._drag = False


def _no_cursor(method):
    """Run a figure-saving method of the application with the bar taken out of the figure, then put it back."""
    import functools

    @functools.wraps(method)
    def wrapper(self, *args, **kwargs):
        fig = self._current_figure()
        cur = self._cursor_of(fig) if fig is not None else None
        if cur is not None:
            cur.remove()
        try:
            return method(self, *args, **kwargs)
        finally:
            if cur is not None:
                cur.attach()
                cur.canvas.draw_idle()
    return wrapper


# ── Main application ──────────────────────────────────────────────────────────

class DLSManager:
    def __init__(self, root):
        self.root = root
        self.root.title("DLS Manager — Horiba SZ-100")
        try:      # the icon next to the program; the window keeps the default one if the file is missing
            self.root.iconbitmap(os.path.join(os.path.dirname(os.path.abspath(__file__)), "DLSManager.ico"))
        except tk.TclError:
            pass
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
        m.add_command(label="Open .nsz / .csv files...", command=self._open_files)
        m.add_command(label="Open folder...", command=self._open_folder)
        m.add_command(label="Average selected measurements", command=self._average_selected)
        m.add_command(label="Export Excel...", command=self._export_xlsx)
        m.add_separator()
        m.add_command(label="Save Figure Image (PNG, PDF, SVG)...", command=self.save_figure_image)
        m.add_command(label="Save Figure (pickle)...", command=self.save_figure_pickle)
        m.add_command(label="Edit Figure...", command=self.open_figure_editor)
        m.add_command(label="PlotStyleKit...", command=self.plotstylekit_dialog)
        m.add_separator()
        m.add_command(label="Exit", command=self.root.destroy)
        bar.add_cascade(label="File", menu=m)
        hm = tk.Menu(bar, tearoff=0)
        hm.add_command(label="SZ-100 settings and recalculation (guide)", command=self.open_guide)
        bar.add_cascade(label="Help", menu=hm)
        self.root.config(menu=bar)

    GUIDE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs", "sz100_guide.html")

    def open_guide(self):
        """Show the guide to the SZ-100 settings and recalculation (docs/sz100_guide.html, next to the program) in a window of
        the program; a second call brings the open window to the front."""
        if not os.path.isfile(self.GUIDE):
            messagebox.showinfo("Guide", "The guide is not found next to the program:\n" + self.GUIDE)
            return
        win = getattr(self, "_guide_win", None)
        if win is not None and win.winfo_exists():
            win.deiconify(); win.lift(); win.focus_force()
            return
        self._guide_win = GuideWindow(self.root, self.GUIDE)

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
            ("＋  Open .nsz / .csv files", self._open_files),
            ("📁  Open folder",  self._open_folder),
            ("∑  Average selected",   self._average_selected),
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
            text="⬇ Drop .nsz / .csv files here",
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

        # Tab 2: Size distribution (the instrument's own result or a new analysis of the correlogram)
        self.tab_size = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.tab_size, text="  Size distribution  ")

        # Tab 3: Parameters
        self.tab_params = tk.Frame(self.notebook, bg="white")
        self.notebook.add(self.tab_params, text="  Parameters  ")

        self._build_acf_tab()
        self._build_size_tab()
        self._build_params_tab()
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

        ttk.Separator(ctrl, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        tk.Label(ctrl, text="Quantity:", bg="white", font=("Helvetica",9)).pack(side=tk.LEFT)
        # g1 is what the instrument software shows and what the Size distribution tab plots; the file stores g2 - 1 = g1^2
        self.acf_quantity = tk.StringVar(value="g1")
        for val, lbl in [("g1", "g1"), ("g2", "g2 - 1")]:
            tk.Radiobutton(ctrl, text=lbl, variable=self.acf_quantity, value=val,
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
        self.cur_acf = PlotCursor(self.canvas_acf)

    def _build_size_tab(self):
        """One tab for the size distribution, drawn the same way for both sources: the correlogram with its fit (g1) on the
        left, the size distribution by intensity, volume or number in the middle and, optionally, the cumulative curve on
        the right. Source 'Instrument': the fit and the distribution that the instrument software stored in the file.
        Source 'Recomputed': a new analysis of the stored correlogram (cumulants, free or log-normal distributions).
        Weighting, the conversion used for volume and number and the x scale apply to both."""
        self._tips = []
        self.dls_vars = {}
        bg, fnt = "white", ("Helvetica", 9)
        top = tk.Frame(self.tab_size, bg=bg, pady=4)
        top.pack(fill=tk.X, padx=14)
        tk.Label(top, text="Source:", bg=bg, font=fnt).pack(side=tk.LEFT)
        self.size_source = tk.StringVar(value=SIZE_SOURCES[0])
        for v in SIZE_SOURCES:
            rb = tk.Radiobutton(top, text=v, variable=self.size_source, value=v, bg=bg, font=fnt,
                                command=self._on_source)
            rb.pack(side=tk.LEFT, padx=2)
            self._tips.append(ToolTip(rb, DLS_TIPS['source']))
        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        tk.Label(top, text="X scale:", bg=bg, font=fnt).pack(side=tk.LEFT)
        self.size_xscale = tk.StringVar(value="log")
        for val, lbl in [("log", "Log"), ("linear", "Linear")]:
            tk.Radiobutton(top, text=lbl, variable=self.size_xscale, value=val, bg=bg, font=fnt,
                           command=self._plot_dist).pack(side=tk.LEFT, padx=4)
        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        tk.Label(top, text="Weighting:", bg=bg, font=fnt).pack(side=tk.LEFT)
        self.dls_weight = tk.StringVar(value=DLS_WEIGHTS[0])
        for w in DLS_WEIGHTS:
            rb = tk.Radiobutton(top, text=w, variable=self.dls_weight, value=w, bg=bg, font=fnt, command=self._plot_dist)
            rb.pack(side=tk.LEFT, padx=2)
            self._tips.append(ToolTip(rb, DLS_TIPS['weight']))
        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)
        self.show_cumul = tk.BooleanVar(value=False)
        tk.Checkbutton(top, text="Show cumulative", variable=self.show_cumul, bg=bg, font=fnt,
                       command=self._plot_dist).pack(side=tk.LEFT)

        conv = tk.Frame(self.tab_size, bg=bg)
        conv.pack(fill=tk.X, padx=14)
        lab0 = tk.Label(conv, text="Conversion to volume / number:", bg=bg, font=fnt)
        lab0.pack(side=tk.LEFT)
        self.dls_conv = tk.StringVar(value=DLS_CONVERSIONS[0])
        cv = ttk.Combobox(conv, textvariable=self.dls_conv, values=DLS_CONVERSIONS, state="readonly", width=9)
        cv.pack(side=tk.LEFT, padx=(4, 0))
        cv.bind("<<ComboboxSelected>>", lambda e: self._on_conv())
        self._tips.extend(ToolTip(w, DLS_TIPS['conv']) for w in (lab0, cv))
        # which fields belong to which conversion: Particle n to Mie, the vesicle parameters to Mie shell
        self._conv_fields = {"Mie": [], "Mie shell": [], "Mie|Mie shell": []}
        lab = tk.Label(conv, text="Particle n (Mie; empty = the file's)", bg=bg, font=fnt)
        lab.pack(side=tk.LEFT, padx=(10, 2))
        self.dls_vars['np'] = tk.StringVar(value='')              # empty = the index stored in the file (1.59 if absent)
        ent = tk.Entry(conv, textvariable=self.dls_vars['np'], width=6)
        ent.pack(side=tk.LEFT)
        ent.bind("<Return>", lambda e: self._plot_dist())
        self._tips.extend(ToolTip(w, DLS_TIPS['pn']) for w in (lab, ent))
        self._conv_fields["Mie"] += [lab, ent]
        lab = tk.Label(conv, text="Aperture ± (deg)", bg=bg, font=fnt)
        lab.pack(side=tk.LEFT, padx=(10, 2))
        self.dls_vars['aperture'] = tk.StringVar(value='%g' % DLS_APERTURE)
        ent = tk.Entry(conv, textvariable=self.dls_vars['aperture'], width=4)
        ent.pack(side=tk.LEFT)
        ent.bind("<Return>", lambda e: self._plot_dist())
        self._tips.extend(ToolTip(w, DLS_TIPS['aperture']) for w in (lab, ent))
        self._conv_fields["Mie|Mie shell"] += [lab, ent]       # used by both Mie conversions
        shell = tk.Frame(self.tab_size, bg=bg)
        shell.pack(fill=tk.X, padx=14)
        lab0 = tk.Label(shell, text="Vesicle (Mie shell):", bg=bg, font=fnt)
        lab0.pack(side=tk.LEFT)
        self._conv_fields["Mie shell"].append(lab0)
        for text, key, val, width in (("lamellae", 'shell_n', '1', 4), ("bilayer (nm)", 'shell_t', '4.5', 5),
                                      ("repeat distance (nm)", 'shell_rep', '7.0', 5),
                                      ("index excess of the lipid", 'shell_dn', '0.12', 5)):
            lab = tk.Label(shell, text=text, bg=bg, font=fnt)
            lab.pack(side=tk.LEFT, padx=(10, 2))
            self.dls_vars[key] = tk.StringVar(value=val)
            ent = tk.Entry(shell, textvariable=self.dls_vars[key], width=width)
            ent.pack(side=tk.LEFT)
            ent.bind("<Return>", lambda e: self._plot_dist())
            self._tips.extend(ToolTip(w, DLS_TIPS['shell']) for w in (lab, ent))
            self._conv_fields["Mie shell"] += [lab, ent]
        self.ray_warn = tk.Label(self.tab_size, text="", bg=bg, fg="#B91C1C", font=("Helvetica", 9, "bold"),
                                 anchor="w", justify=tk.LEFT, wraplength=1000)
        self.ray_warn.pack(fill=tk.X, padx=14)
        self._on_conv(draw=False)

        # controls of the recomputed analysis (shown only for that source), above the one plot area
        self.box_rec = tk.Frame(self.tab_size, bg=bg)
        ctrl = tk.Frame(self.box_rec, bg=bg)
        ctrl.pack(fill=tk.X)
        tk.Label(ctrl, text="Model:", bg=bg, font=fnt).pack(side=tk.LEFT)
        self.dls_model = tk.StringVar(value=DLS_MODELS[0])
        cmb = ttk.Combobox(ctrl, textvariable=self.dls_model, values=DLS_MODELS, state="readonly", width=30)
        cmb.pack(side=tk.LEFT, padx=(4, 0))
        self._tips.append(ToolTip(cmb, DLS_TIPS['model']))
        for text, key, val in (("Wavelength (nm)", 'wl', '532'), ("Smoothing α", 'alpha', '1.0'),
                               ("g1 min", 'gmin', '0.1')):
            lab = tk.Label(ctrl, text=text, bg=bg, font=fnt)
            lab.pack(side=tk.LEFT, padx=(12, 2))
            self.dls_vars[key] = tk.StringVar(value=val)
            ent = tk.Entry(ctrl, textvariable=self.dls_vars[key], width=6)
            ent.pack(side=tk.LEFT)
            ent.bind("<Return>", lambda e: self._plot_dist())
            self._tips.extend(ToolTip(w, DLS_TIPS[key]) for w in (lab, ent))
        # the usual g1 cut-off depends on the model (0.1 regularised, 0.05 peaks)
        self.dls_model.trace_add("write", self._on_dls_model)
        ctrl2 = tk.Frame(self.box_rec, bg=bg)
        ctrl2.pack(fill=tk.X)
        self.dls_range = tk.BooleanVar(value=True)
        rb = tk.Checkbutton(ctrl2, text="Instrument range", variable=self.dls_range, bg=bg, font=fnt,
                            command=self._plot_dist)
        rb.pack(side=tk.LEFT)
        self._tips.append(ToolTip(rb, DLS_TIPS['range']))
        lab = tk.Label(ctrl2, text="Tail cut (% of max)", bg=bg, font=fnt)
        lab.pack(side=tk.LEFT, padx=(12, 2))
        self.dls_vars['tail'] = tk.StringVar(value='5')
        ent = tk.Entry(ctrl2, textvariable=self.dls_vars['tail'], width=4)
        ent.pack(side=tk.LEFT)
        ent.bind("<Return>", lambda e: self._plot_dist())
        self._tips.extend(ToolTip(w, DLS_TIPS['tail']) for w in (lab, ent))
        lab = tk.Label(ctrl2, text="Width (instrument form)", bg=bg, font=fnt)
        lab.pack(side=tk.LEFT, padx=(12, 2))
        self.dls_width = tk.StringVar(value=DLS_WIDTHS[0])
        cmb = ttk.Combobox(ctrl2, textvariable=self.dls_width, values=list(DLS_WIDTHS), state="readonly", width=11)
        cmb.pack(side=tk.LEFT)
        cmb.bind("<<ComboboxSelected>>", lambda e: self._on_width())
        self._tips.extend(ToolTip(w, DLS_TIPS['width']) for w in (lab, cmb))
        lab = tk.Label(ctrl2, text="Resolution", bg=bg, font=fnt, fg="#94A3B8")
        lab.pack(side=tk.LEFT, padx=(8, 2))
        self.dls_vars['resolution'] = tk.StringVar(value='0.3')
        self.ent_resolution = tk.Entry(ctrl2, textvariable=self.dls_vars['resolution'], width=5, state=tk.DISABLED)
        self.ent_resolution.pack(side=tk.LEFT)
        self.ent_resolution.bind("<Return>", lambda e: self._plot_dist())
        self.lab_resolution = lab
        self._tips.extend(ToolTip(w, DLS_TIPS['resolution']) for w in (lab, self.ent_resolution))
        btn = tk.Button(ctrl2, text="Compute", command=self._plot_dist, bg="#FEF9C3", font=fnt)
        btn.pack(side=tk.LEFT, padx=12)

        # the plot area: text with the values at the bottom, then the toolbar, then the figure
        self.plot_area = tk.Frame(self.tab_size, bg=bg)
        self.plot_area.pack(fill=tk.BOTH, expand=True)
        self.dls_text = tk.Text(self.plot_area, height=9, font=("Consolas", 9), bg="#F8FAFC", relief=tk.FLAT,
                                wrap=tk.WORD, state=tk.DISABLED)
        self.dls_text.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 6))
        self.fig_size = Figure(figsize=(10, 4.5), dpi=100, facecolor="white")
        self.canvas_size = FigureCanvasTkAgg(self.fig_size, self.plot_area)
        NavigationToolbar2Tk(self.canvas_size, self.plot_area).pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas_size.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=6, pady=4)
        self.cur_size = PlotCursor(self.canvas_size)
        self._on_source()

    def _on_conv(self, draw=True):
        """Enable the fields of the chosen conversion and grey out the others, then redraw."""
        conv = self.dls_conv.get()
        for name, widgets in self._conv_fields.items():
            on = conv in name.split("|")
            for w in widgets:
                if isinstance(w, tk.Entry):
                    w.config(state=tk.NORMAL if on else tk.DISABLED)
                else:
                    w.config(fg="black" if on else "#94A3B8")
        if draw:
            self._plot_dist()

    def _on_source(self):
        """Show the controls of the chosen source and draw it."""
        self.box_rec.pack_forget()
        if self.size_source.get() == SIZE_SOURCES[1]:
            self.box_rec.pack(fill=tk.X, padx=14, before=self.plot_area)
        self._plot_dist()

    def _build_params_tab(self):
        frame = tk.Frame(self.tab_params, bg="white")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        cols = ("Parameter", "Value", "Where it comes from")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.heading("Parameter", text="Parameter")
        self.tree.column("Parameter", width=170, anchor="w", stretch=False)
        self.tree.heading("Value", text="Value")
        self.tree.column("Value", width=190, anchor="center", stretch=False)
        self.tree.heading("Where it comes from", text="Where it comes from")
        self.tree.column("Where it comes from", width=720, anchor="w")

        style = ttk.Style()
        style.configure("Treeview", font=("Helvetica",9), rowheight=22)
        style.configure("Treeview.Heading", font=("Helvetica",9,"bold"))
        style.map("Treeview",
                  background=[("selected","#DBEAFE")],
                  foreground=[("selected","#1E3A5F")])

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Alternating rows
        self.tree.tag_configure("odd",  background="#F8FAFC")
        self.tree.tag_configure("even", background="white")

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
            if os.path.isfile(p) and p.lower().endswith(('.nsz', '.csv')):
                nsz_files.append(p)
            elif os.path.isdir(p):
                for f in os.listdir(p):
                    if f.lower().endswith('.nsz'):
                        nsz_files.append(os.path.join(p, f))
        self._load_files(nsz_files)

    # ── File loading ──────────────────────────────────────────────────────────

    def _open_files(self):
        files = filedialog.askopenfilenames(
            title="Open NSZ files or CSV exports",
            filetypes=[("Horiba NSZ and CSV exports", "*.nsz *.csv"), ("Horiba NSZ", "*.nsz"), ("CSV export", "*.csv"),
                       ("All Files", "*.*")])
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
                data = load_csv(path) if path.lower().endswith('.csv') else load_nsz(path)
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

    def _average_selected(self):
        """Add the mean of the selected measurements to the list and select it (see dls_average)."""
        if len(self.selected) < 2:
            messagebox.showinfo("Average", "Select two or more measurements in the list (Ctrl+click or Shift+click).")
            return
        try:
            data = dls_average([self.samples[p] for p in self.selected])
        except ValueError as e:
            messagebox.showwarning("Average", str(e))
            return
        path, stem, k = data['path'], os.path.splitext(data['path'])[0], 2
        while path in self.samples and self.samples[path]['sources'] != data['sources']:      # same name, other files
            path, k = f"{stem}_{k}.avg", k + 1
        if path not in self.samples:
            data['path'], data['name'] = path, os.path.splitext(os.path.basename(path))[0]
            self.samples[path] = data
            self.listbox.insert(tk.END, "∑ " + data['name'])
        index = list(self.samples).index(path)
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(index)
        self.listbox.see(index)
        self._on_select()

    def _data_folder(self):
        """Where files are saved by default: the folder of the input file of the first selected sample (of the first of its
        measurements for an average), else of the last one loaded; None with nothing loaded."""
        path = self.selected[0] if self.selected else next(reversed(self.samples), None)
        return os.path.dirname(os.path.abspath(path)) if path else None

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
        if self.selected:
            matplotlib.rcParams['savefig.directory'] = self._data_folder()      # the toolbar's own save button
        self._refresh_all()

    def _refresh_all(self):
        self._plot_acf()
        self._plot_dist()
        self._update_params_table()

    def _refresh_dls(self):
        """Redraw the size distribution tab (the recomputed source only while the tab is shown: see _plot_dist)."""
        self._plot_dist()

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
        g1_view = self.acf_quantity.get() == "g1"
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

            conv_q = (lambda v: np.sqrt(np.clip(v, 0, None))) if g1_view else (lambda v: v)
            ax.plot(tau[mask], conv_q(acf[mask]), 'o', color=c, markersize=3.5,
                    linewidth=0, label=lbl)

            # Ideal fit — aligned to tau[fit_offset:]
            off = d.get('fit_offset', 6)
            if self.show_fit.get() and len(d['fit']) > 2:
                tau_fit = tau[off:off + len(d['fit'])]
                fit_vals = conv_q(d['fit'][:len(tau_fit)])
                ax.plot(tau_fit, fit_vals, '-', color=c, linewidth=1.5,
                        alpha=0.85)[0]._cursor_name = lbl + " fit"

            # Residuals — same offset as the fit
            if axr is not None and len(d['resid']) > 2:
                tau_res = tau[off:off + len(d['resid'])]
                res_vals = d['resid'][:len(tau_res)]
                axr.plot(tau_res, res_vals, 'o', color=c, markersize=2.5,
                         linewidth=1, alpha=0.8)

        ax.set_xscale(xscale)
        ax.set_xlabel("Delay time τ (µs)", fontsize=10)
        # mathtext because Arial has no subscript glyphs; the stream holds g1 squared (see CLAUDE.md)
        ax.set_ylabel(r"$g_1(\tau)$ = √(g$_2$ − 1), normalized" if g1_view else r"$g_2(\tau) - 1$ normalized", fontsize=10)
        ax.set_ylim(-0.05, 1.1)
        ax.set_title("Autocorrelation function (ACF)", fontsize=11, fontweight='bold')
        ax.axhline(0, color='gray', linewidth=0.7, linestyle='--')
        if self.selected:
            ax.legend(fontsize=8, loc='upper right', framealpha=0.9)

        if axr is not None:
            axr.set_xscale(xscale)
            axr.set_xlabel("Delay time τ (µs)", fontsize=10)
            axr.set_ylabel("Residuals (g$_2$ − 1)", fontsize=9)
            axr.axhline(0, color='gray', linewidth=0.8)
            plt.setp(ax.get_xticklabels(), visible=False)
            ax.set_xlabel("")

        if not self.selected:
            ax.text(0.5, 0.5, "No sample selected\n\nLoad .nsz files and select them in the list",
                    ha='center', va='center', transform=ax.transAxes,
                    fontsize=12, color='#94A3B8')

        self._style_figure(self.fig_acf)
        self.cur_acf.attach()
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

    def _plot_size(self):
        """Kept for the callers of the former instrument-only plot: the one size distribution plot."""
        self._plot_dist()

    # ── DLS analysis ──────────────────────────────────────────────────────────

    def _on_dls_model(self, *_):
        self.dls_vars['gmin'].set('0.1' if self.dls_model.get() == DLS_MODELS[0] else '0.05')
        self._plot_dist()

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
        """Kept for the callers of the former recomputed-only plot: the one size distribution plot."""
        self._plot_dist()

    def _inst_summary(self, name, d):
        """Text lines for one sample: the values that the instrument software stored in the file."""
        p = d['params']

        def f(key, fmt):
            return fmt.format(p[key]) if p.get(key) is not None else "n/a"

        return [name,
                "  Instrument: Z-average %s nm, PI %s | mean %s nm, mode %s nm, S.D. %s nm | D10/D50/D90 %s / %s / %s nm"
                % (f('ZAvg_nm', '{:.1f}'), f('PDI', '{:.3f}'), f('Mean1_nm', '{:.0f}'), f('Mode_nm', '{:.0f}'),
                   f('SD1_nm', '{:.0f}'), f('D10_nm', '{:.0f}'), f('D50_nm', '{:.0f}'), f('D90_nm', '{:.0f}'))]

    def _plot_dist(self):
        """The size distribution tab, the same layout for both sources: correlogram (g1) with the fit, size distribution,
        and optionally the cumulative curve. Instrument: fit and distribution stored in the file. Recomputed: a new
        analysis of the stored correlogram. The recomputed analysis takes a moment, so it is drawn only while the tab is
        shown (<<NotebookTabChanged>> calls this when the tab is opened)."""
        inst = self.size_source.get() == SIZE_SOURCES[0]
        if not inst and self.notebook.select() != str(self.tab_size):
            return
        fig = self.fig_size
        fig.clear()
        weight, conv = self.dls_weight.get(), self.dls_conv.get()
        show_c = self.show_cumul.get()
        # no wspace here: a gridspec with its own spacing is left out of tight_layout, and the saved figure had its x labels cut
        gs = fig.add_gridspec(1, 3 if show_c else 2)
        ax1, ax2 = fig.add_subplot(gs[0]), fig.add_subplot(gs[1])
        axc = fig.add_subplot(gs[2]) if show_c else None
        lines, ray_items, peak_warn = [], [], []
        wl = alpha = gmin = None
        model = self.dls_model.get()
        ok = True
        if not inst:
            try:
                wl, alpha, gmin = (float(self.dls_vars[k].get()) for k in ('wl', 'alpha', 'gmin'))
                if wl <= 0 or alpha < 0 or not 0 < gmin < 1:
                    raise ValueError
            except ValueError:
                ok = False
                lines.append("Invalid settings: wavelength > 0, smoothing >= 0 and 0 < g1 min < 1 are required.")
            try:
                import scipy.optimize  # noqa: F401  (nnls and least_squares)
            except ImportError:
                ok = False
                lines.append("The inversion needs scipy (pip install scipy).")
        palette = sample_colors()
        if not inst and ok and self.selected:
            self.root.config(cursor="watch")
            self.root.update_idletasks()
        try:
            for idx, path in enumerate(self.selected[:MAX_SAMPLES_SHOWN] if ok else []):
                d = self.samples[path]
                c = palette[idx % len(palette)]
                name = self._short_label(d['name'])
                rec = dls_record(d)
                r = None
                if not inst:
                    if rec is None:
                        lines.append(f"{name}: the file has no correlogram or measurement conditions to analyse.")
                        continue
                    try:
                        r = self._dls_result(path, model, wl, alpha, gmin)
                    except Exception as e:
                        lines.append(f"{name}: inversion failed ({e}).")
                        continue
                # ---- correlogram as g1 = sqrt(g2 - 1), with the instrument's fit or the recomputed one
                if d.get('tau_aligned'):
                    tau = d['tau']
                    g1 = np.sqrt(np.clip(d['acf'], 0, None))
                    rng = None
                    if inst and d['params'].get('Range_from'):
                        rng = (d['params']['Range_from'] - 1, d['params']['Range_to'])
                    elif not inst and self.dls_range.get() and rec.get('range'):
                        rng = rec['range']
                    if rng:
                        out_of = np.ones(len(g1), bool)
                        out_of[rng[0]:rng[1]] = False
                        ax1.plot(tau[out_of], g1[out_of], 'o', ms=3, mfc='none', color=c, alpha=0.5)
                        ax1.plot(tau[~out_of], g1[~out_of], 'o', ms=3, color=c, label=name)
                    else:
                        ax1.plot(tau, g1, 'o', ms=3, color=c, label=name)
                    if inst:
                        off = d.get('fit_offset', 6)
                        if len(d['fit']) > 2:
                            ft = tau[off:off + len(d['fit'])]
                            ax1.plot(ft, np.sqrt(np.clip(d['fit'][:len(ft)], 0, None)), '-', lw=1.3,
                                     color=c)[0]._cursor_name = name + " fit"
                    else:
                        # the line only spans the delays that the inversion used: g1 above its minimum, delays of at least
                        # 1 us and, with the option, the instrument's channels (not the whole axis)
                        used = (g1 > gmin) & (tau >= 1.0)
                        if rng:
                            used &= ~out_of
                        idx = np.nonzero(used)[0]
                        if len(idx) > 1:
                            sl = slice(idx[0], idx[-1] + 1)
                            ax1.plot(tau[sl], r['fit'][sl], '-', lw=1.3, color=c)[0]._cursor_name = name + " fit"
                # ---- size distribution
                if inst:
                    sz = d['sz']
                    m = sz > 0
                    if m.sum() < 2:
                        continue
                    x, w0 = sz[m], d['intens'][m]
                    if weight == DLS_WEIGHTS[0]:
                        w, cum = w0, d['cumul'][m]
                    else:
                        i1 = np.ones(len(x))
                        live = w0 > 0
                        if rec is not None and conv != "Rayleigh" and live.any():
                            i1[live] = self._dls_i1(d, rec, x[live], conv, rec_wl(d))
                        else:
                            i1 = x ** 6
                        w = dls_reweight(x, w0, weight, i1)
                        cum = np.cumsum(w) / w.sum() * 100.0 if w.sum() > 0 else w
                    ray_items.append((name, x, w0, rec, rec_wl(d), self._particle_index(d)))
                    z = d['params'].get('ZAvg_nm')
                    lines.extend(self._inst_summary(name, d))
                    ax2.plot(x, w / w.max(), 'o-', color=c, ms=2.5, lw=1.4, label=name)
                else:
                    i1 = self._dls_i1(d, rec, r['d'], conv, wl) if weight != DLS_WEIGHTS[0] else None
                    pw = self._weighted(r, weight, i1)
                    x, w, cum = r['d'], pw, np.cumsum(pw) / pw.sum() * 100.0
                    ray_items.append((name, r['d'], r['p'], rec, wl, self._particle_index(d)))
                    peak_warn += self._peak_warnings(path, name, r, wl, alpha, gmin)
                    z = r['z']
                    lines.extend(self._dls_summary(path, name, r, model, wl, alpha, gmin, weight, i1, conv))
                    ax2.plot(x, pw / pw.max(), 'o-', color=c, ms=2.5, lw=1.4, label=name)    # same style as the Instrument source
                if z:
                    ax2.axvline(z, color=c, ls=':', lw=1)
                if axc is not None:
                    axc.plot(x, cum, '-', color=c, lw=1.5, label=name)
                    if inst and weight == DLS_WEIGHTS[0]:
                        for pct in (10, 50, 90):
                            val = d['params'].get(f'D{pct}_nm')
                            if val:
                                axc.axvline(val, color=c, lw=0.7, ls='--', alpha=0.6)
        finally:
            self.root.config(cursor="")
        xscale = self.size_xscale.get()
        ax1.set_xscale('log')
        ax1.set_ylim(-0.05, 1.2)                    # headroom for the legend
        ax2.set_ylim(-0.02, 1.35)
        ax1.set_xlabel("Delay time τ (µs)")
        ax1.set_ylabel("$g_1$ = √ACF")
        ax1.set_title("Correlogram and fit", fontweight='bold')
        ax2.set_xscale(xscale)
        # the same diameter axis for both sources: the instrument's table starts at 0.34 nm and the recomputed grid at
        # 10 nm, and nothing below 10 nm is measurable here; up to 1e4 nm (the table ends at 8510)
        size_xlim = (10.0, 1e4) if xscale == 'log' else (0.0, 1000.0)
        ax2.set_xlim(*size_xlim)
        ax2.set_xlabel("Diameter (nm)")
        ax2.set_ylabel(f"{weight}-weighted (normalised)")
        ax2.set_title(f"Size distribution by {weight.lower()}", fontweight='bold')
        if self.selected and ax1.get_legend_handles_labels()[0]:
            ax1.legend(fontsize=8, loc='upper right', framealpha=0.9)
        if self.selected and any(ln.get_linestyle() == ':' for ln in ax2.lines):
            ax2.plot([], [], ':', color='gray', lw=1,
                     label='Z-average (instrument)' if inst else 'Z-average (cumulants)')
            ax2.legend(fontsize=8, loc='upper right', framealpha=0.9)
        if axc is not None:
            axc.set_xscale(xscale)
            axc.set_xlim(*size_xlim)
            axc.set_xlabel("Diameter (nm)")
            axc.set_ylabel("Cumulative (%)")
            axc.set_title(f"Cumulative distribution by {weight.lower()}", fontweight='bold')
            axc.set_ylim(-2, 104)
            for pct in (10, 50, 90):
                axc.axhline(pct, color='gray', linewidth=0.5, linestyle=':')
            if self.selected:
                if any(ln.get_linestyle() == '--' for ln in axc.lines):
                    axc.plot([], [], '--', color='gray', lw=0.8, label='D10, D50, D90')
                axc.legend(fontsize=8, loc='upper left', framealpha=0.9)
        if not self.selected:
            ax1.text(0.5, 0.5, "No sample selected\n\nLoad .nsz files and select them in the list",
                     ha='center', va='center', transform=ax1.transAxes, fontsize=12, color='#94A3B8')
        self._show_warnings([self._volnum_warning(weight), self._rayleigh_warning(ray_items, weight, conv)] + peak_warn)
        self._set_dls_text(lines)
        self._style_figure(fig)
        # PlotStyleKit's style() autoscales x tight to the data, which undoes the limits set above: set them again
        ax2.set_xlim(*size_xlim)
        if axc is not None:
            axc.set_xlim(*size_xlim)
        self.cur_size.attach()
        self.canvas_size.draw()

    def _rayleigh_limit(self, rec, wl, n_particle):
        """Diameter (nm) above which the Rayleigh d^6 law is off by more than 10% for these conditions: the first size at
        which Mie scattering of a solid sphere, divided by d^6, falls below 0.9 of its small-size value. 60 nm if the
        file has no angle or index to work it out."""
        if rec is None:
            return 60.0
        key = ('ray', rec['angle'], rec['n'], wl, n_particle)
        if key not in self._dls_cache:
            grid = np.geomspace(5.0, 400.0, 70)
            ratio = mie_intensity(grid, wl, rec['n'], n_particle, rec['angle']) / grid ** 6
            below = np.nonzero(ratio < 0.9 * ratio[0])[0]
            self._dls_cache[key] = float(grid[below[0]]) if len(below) else 400.0
        return self._dls_cache[key]

    def _rayleigh_warning(self, items, weight, conv):
        """Red note above the plot when Rayleigh is used for volume or number on a distribution with much of its
        intensity above the size where d^6 stops being right. `items`: (name, diameters, intensity weights, record,
        wavelength, particle index) of the plotted samples."""
        text = ""
        if conv == "Rayleigh" and weight != DLS_WEIGHTS[0]:
            worst = None
            for name, d, p, rec, wl, n_p in items:
                lim = self._rayleigh_limit(rec, wl, n_p)
                p = np.asarray(p, float)
                frac = float(p[np.asarray(d, float) > lim].sum() / p.sum()) if p.sum() > 0 else 0.0
                if frac > 0.05 and (worst is None or frac > worst[0]):
                    worst = (frac, lim, name)
            if worst:
                text = ("Warning: Rayleigh is more than 10%% off above %.0f nm and %.0f%% of the intensity of %s is larger: "
                        "volume and number are not reliable, use Mie." % (worst[1], 100 * worst[0], worst[2]))
        return text

    @staticmethod
    def _volnum_warning(weight):
        """Always-visible caution for volume and number."""
        if weight == DLS_WEIGHTS[0]:
            return ""
        return ("Caution: %s is not measured. It is the intensity distribution divided by the light one particle scatters, "
                "which amplifies its small-size side and its width, so the values are only indicative; trust the "
                "intensity distribution and the Z-average." % weight.lower())

    def _peak_warnings(self, path, name, r, wl, alpha, gmin):
        """Cautions about the log-normal fits of a recomputed sample: a peak at the width limit of the fit, a second peak
        that the BIC does not support, two peaks closer than a factor 3 (the split is ill-posed)."""
        out = []
        if r['peaks'] is None:
            return out
        if any(pk['sigma'] >= 1.15 for pk in r['peaks']):
            out.append("%s: a peak reaches the width limit of the fit (sigma 1.2): it is a tail, not a population." % name)
        if len(r['peaks']) == 2:
            r1 = self._dls_result(path, DLS_MODELS[1], wl, alpha, gmin)
            if r['bic'] > r1['bic'] - 6:
                out.append("%s: the second peak is not supported (BIC): one peak describes the data as well." % name)
            elif r['peaks'][1]['median'] < 3 * r['peaks'][0]['median']:
                out.append("%s: the two peaks are closer than a factor 3: the split is not stable between replicates." % name)
        return out

    def _show_warnings(self, parts):
        self.ray_warn.config(text="\n".join(x for x in parts if x))

    def _tail_fraction(self):
        try:
            return min(max(float(self.dls_vars['tail'].get()), 0.0), 50.0) / 100.0
        except ValueError:
            return 0.05

    def _weighted(self, r, weight, i1):
        """Distribution `r` of dls_analyse() in the chosen weighting (the tail cut applies to the free model)."""
        p = r['p']
        if r['peaks'] is None:
            p = self._with_width(r['d'], p)                          # every weighting, the intensity one included
            if weight != DLS_WEIGHTS[0]:
                p = dls_floor(p, self._tail_fraction())
        return dls_reweight(r['d'], p, weight, i1)

    def _width_sigma(self):
        """sigma of ln d to give the free distribution: the one of the instrument's form chosen in the tab, or of the typed
        resolution (Manual; invalid -> the Standard 0.3, limited to -1..1); None for 'As computed'."""
        form = self.dls_width.get()
        if form == "Manual":
            try:
                res = min(max(float(self.dls_vars['resolution'].get()), -1.0), 1.0)
            except ValueError:
                res = DLS_FORM_RESOLUTION["Standard"]
            return dls_resolution_sigma(res)
        if form in DLS_FORM_RESOLUTION:
            return dls_resolution_sigma(DLS_FORM_RESOLUTION[form])
        return None

    def _with_width(self, d, p):
        """The free distribution `p` on the grid `d` with the width chosen in the tab."""
        s = self._width_sigma()
        return p / p.sum() if s is None else dls_set_width(d, p, s)

    def _width_note(self):
        """Text for the plot summary when the width was set ('' for 'As computed')."""
        s = self._width_sigma()
        if s is None:
            return ""
        return f", width set to σ(ln d) {s:.2f} as the instrument's {self.dls_width.get()} (assumption)"

    def _on_width(self):
        """Enable the resolution entry only for the Manual form, then redraw."""
        manual = self.dls_width.get() == "Manual"
        self.ent_resolution.config(state=tk.NORMAL if manual else tk.DISABLED)
        self.lab_resolution.config(fg="black" if manual else "#94A3B8")
        self._plot_dist()

    def _particle_index(self, d):
        """Refractive index of the particle: the value typed in the tab, else the one in the file, else 1.59."""
        try:
            v = float(self.dls_vars['np'].get())
            if v > 1.0:
                return v
        except ValueError:
            pass
        return d['params'].get('n_particle') or 1.59

    def _shell_params(self):
        """(lamellae, bilayer nm, repeat distance nm, index excess of the lipid) typed for the vesicle; the defaults of
        LabSpectrumManager (one bilayer of 4.5 nm, repeat 7 nm, 0.12) for any value that is invalid."""
        try:
            n, t = int(float(self.dls_vars['shell_n'].get())), float(self.dls_vars['shell_t'].get())
            rep_, dn = float(self.dls_vars['shell_rep'].get()), float(self.dls_vars['shell_dn'].get())
            if 1 <= n <= 60 and t > 0 and rep_ >= t and 0 < dn < 1:
                return n, t, rep_, dn
        except ValueError:
            pass
        return 1, 4.5, 7.0, 0.12

    def _shell(self):
        """(thickness nm, index excess) of the one shell that stands for the vesicle: n lamellae are a shell of
        n * repeat - (repeat - bilayer) whose index excess is that of the lipid averaged over the water gaps (the model of
        LabSpectrumManager's vesicle_amplitude); one lamella is the bilayer itself."""
        n, t, rep_, dn = self._shell_params()
        total = n * rep_ - (rep_ - t)
        return total, dn * n * t / total

    def _dls_i1(self, d, rec, grid, conv, wl):
        """Light one particle scatters at each diameter of `grid` (None for Rayleigh: dls_reweight then uses d^6)."""
        if conv == "Rayleigh":
            return None
        key = ('i1', d['path'], conv, wl, self._particle_index(d), self._shell(), len(grid), self._aperture())
        if key not in self._dls_cache:
            self._dls_cache[key] = dls_per_particle(grid, conv, rec, wl, self._particle_index(d), self._shell(),
                                                    self._aperture())
        return self._dls_cache[key]

    def _aperture(self):
        """Half-width (degrees) of the detector aperture typed in the tab, limited to 0-45 (a value above 45 used to fall
        back silently to the default, which looked like the result jumping back); the default DLS_APERTURE if not a number."""
        try:
            return min(max(float(self.dls_vars['aperture'].get()), 0.0), 45.0)
        except ValueError:
            return DLS_APERTURE

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
        apt = f", aperture ±{self._aperture():g}°" if self._aperture() > 0 else ""
        how = {"Mie": f"Mie, particle n {self._particle_index(self.samples[path]):g}{apt}",
               "Mie shell": "Mie shell, %d lamella%s (shell %.3g nm, index excess %.3g%s)" % (
                   self._shell_params()[0], "e" if self._shell_params()[0] > 1 else "", *self._shell(), apt)}.get(conv, "Rayleigh (d^6)")
        if r['peaks'] is None:
            mean, mode, sd = dls_stats(r['d'], self._weighted(r, weight, i1))
            narrowed = self._width_note()
            out.append(f"  Regularised distribution by {w} (α = {alpha:g}{'' if by_int else f', tail cut {100 * self._tail_fraction():g}%'}{narrowed}): mean {mean:.0f} nm, mode {mode:.0f} nm, "
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

    # (label, key of the parameters, where the value comes from); a None key makes a section heading
    PARAM_ROWS = [
        ("── Measurement ──", None, ""),
        ("Date / Time", None, ""),                      # filled from two keys, see below
        ("Angle (°)", 'Angle_deg', "Entered: detector position (90° or 173°); it enters every size"),
        ("Wavelength (nm)", 'Lambda_nm', "Instrument constant: the laser (532 nm)"),
        ("Temperature (°C)", 'Temp_C', "Measured: temperature sensor of the sample holder"),
        ("Viscosity (mPas)", 'Visc_mPas', "Entered: viscosity of the dispersion medium; it enters every size"),
        ("Solvent n", 'n_solvent', "Entered: refractive index of the dispersion medium"),
        ("Solvent", 'Solvent', "Entered: name of the dispersion medium chosen in the software"),
        ("Form of distribution", 'FORM', "Entered: form chosen in the software (dispersity, width); it fixes the histogram's width"),
        ("Resolution", 'Resolution', "Entered: the number behind the width (Monodisperse: Broad -0.5, Standard 0.3, Narrow 0.6); only in the .nsz"),
        ("── Cumulants ──", None, ""),
        ("Z-Average (nm)", 'ZAvg_nm', "Cumulants: mean size weighted by intensity, from the decay rate of the correlogram"),
        ("PDI", 'PDI', "Cumulants: relative width of the decay rates (the PI): the polydispersity of the sample"),
        ("Cumulants S.D. (nm)", 'CumSD_nm', "Cumulants: Z-average x sqrt(PI), exactly; not a separate result"),
        ("── Histogram ──", None, ""),
        ("Peak 1 (nm)", 'Peak1_nm', "Histogram: size of the highest bin (= the mode). Its width is set by the form, not by the PDI"),
        ("Mean 1 (nm)", 'Mean1_nm', "Histogram: mean of the size table; its position is near the Z-average"),
        ("SD 1 (nm)", 'SD1_nm', "Histogram: S.D. of the size table, about a quarter of its mean (set by the form)"),
        ("Area 1 (%)", 'Area1_pct', "Histogram: share of the intensity in peak 1"),
        ("Peak 2 (nm)", 'Peak2_nm', "Histogram: second peak, only when the software finds two (Polydisperse form)"),
        ("Mean 2 (nm)", 'Mean2_nm', "Histogram: mean of the second peak"),
        ("D10 (nm)", 'D10_nm', "Histogram: size below which 10% of the intensity lies (cumulative curve)"),
        ("D50 (nm)", 'D50_nm', "Histogram: size below which 50% of the intensity lies (cumulative curve)"),
        ("D90 (nm)", 'D90_nm', "Histogram: size below which 90% of the intensity lies (cumulative curve)"),
        ("Mode (nm)", 'Mode_nm', "Histogram: size of the highest bin"),
        ("Median (nm)", 'Median_nm', "Histogram: size at 50% of the cumulative curve, interpolated"),
        ("Span", 'Span', "Histogram: width of the cumulative curve relative to its centre"),
        ("Arithmetic mean (nm)", 'AriMean_nm', "Histogram: arithmetic mean of the size table"),
        ("Geometric mean (nm)", 'GeoMean_nm', "Histogram: geometric mean of the size table"),
    ]

    def _update_params_table(self):
        """The parameters of the first selected sample, each with where it comes from (measured, entered, cumulants or histogram)."""
        self.tree.delete(*self.tree.get_children())
        if not self.selected:
            self.tree.heading("Value", text="Value")
            return
        params = self.samples[self.selected[0]]['params']
        self.tree.heading("Value", text=params.get('Sample_Name', '') or "Value")
        self.tree.tag_configure("section", background="#EFF6FF", font=("Helvetica", 9, "bold"))
        sample = self.samples[self.selected[0]]
        sources, spread = sample.get('sources'), sample.get('spread', {})
        if sources:
            self.tree.insert("", tk.END, values=("── Average ──", "", ""), tags=("section",))
            self.tree.insert("", tk.END, values=("Average of", f"{len(sources)} measurements",
                                                 "Results: mean ± S.D. between the measurements; curves: mean point by point"),
                             tags=("even",))
            for i, src in enumerate(sources, 1):
                self.tree.insert("", tk.END, values=(f"File {i}", "", os.path.basename(src)), tags=("odd" if i % 2 else "even",))
        for row_i, (label, key, origin) in enumerate(self.PARAM_ROWS):
            if sources and key == 'CumSD_nm':
                origin = "Cumulants: mean of the Z-average x sqrt(PI) of each measurement"
            if label.startswith("──"):
                self.tree.insert("", tk.END, values=(label, "", ""), tags=("section",))
                continue
            if label == "Date / Time":
                v, origin = f"{params.get('Date','')} {params.get('Time','')}", "Stored by the software when the measurement was made"
            elif key == 'FORM':
                v = params.get('FormText') or dls_form_name(params.get('Dispersity'), params.get('Resolution'))
            else:
                v = params.get(key)
            if v is not None and key in spread:
                v = f"{v} ± {spread[key]:.3g}"
            self.tree.insert("", tk.END, values=(label, str(v) if v is not None else "—", origin),
                             tags=("odd" if row_i % 2 else "even",))


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
            if key in d.get('spread', {}):
                c3 = ws_p.cell(ri, 3, round(d['spread'][key], 4))
                c3.font = cell_font; c3.border = border
            ri += 1

        if d.get('sources'):                                    # an average: say so, and of which files
            style_header(ws_p.cell(1, 3), 'S.D. between measurements')
            ws_p.column_dimensions['C'].width = 28
            for i, src in enumerate(d['sources'], 1):
                ws_p.cell(ri, 1, f'Averaged file {i}').font = cell_font
                ws_p.cell(ri, 2, os.path.basename(src)).font = cell_font
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
                initialdir=self._data_folder(),
                initialfile=f'{d["name"]}.xlsx',
                defaultextension='.xlsx',
                filetypes=[("Excel", "*.xlsx")])
            if dest:
                self._write_xlsx(dest, d)
                messagebox.showinfo("Export completed", f"Saved:\n{dest}")
        else:
            folder = filedialog.askdirectory(title="Destination folder", initialdir=self._data_folder())
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
        return None

    def _cursor_of(self, fig):
        """The PlotCursor of one of the three figures."""
        return {id(self.fig_acf): self.cur_acf, id(self.fig_size): self.cur_size}.get(id(fig))

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
            messagebox.showwarning(title, "Select samples and open the Autocorrelogram or Size distribution tab.")
        return fig

    @_no_cursor
    def save_figure_image(self):
        fig = self._figure_or_warn("Save Figure Image")
        if fig is None:
            return
        if psk is not None:
            psk.save_image_dialog(fig, self.root, initialdir=self._data_folder(), preset=self._preset_for(fig), unpin=False)
            return
        path = filedialog.asksaveasfilename(
            title="Save Figure Image", defaultextension=".png", initialdir=self._data_folder(),
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

    @_no_cursor
    def save_figure_pickle(self):
        """Save the figure as a matplotlib pickle: reopenable with plot_editor.pyw as a live Figure object."""
        fig = self._figure_or_warn("Save Figure")
        if fig is None:
            return
        if psk is not None:
            psk.save_pickle_dialog(fig, self.root, initialdir=self._data_folder(), preset=self._preset_for(fig), unpin=False)
            return
        path = filedialog.asksaveasfilename(
            title="Save Figure", defaultextension=".fig.pickle", initialdir=self._data_folder(),
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

    @_no_cursor
    def open_figure_editor(self):
        """Open the PlotStyleKit Plot Editor on a copy of the current figure."""
        fig = self._figure_or_warn("Edit Figure")
        if fig is None:
            return
        if not self._need_plotstylekit("Edit Figure"):
            return
        editor = psk.edit_dialog(fig, self.root, title="current figure", preset=self._preset_for(fig), unpin=False)
        if editor is not None:
            editor.initialdir = self._data_folder()          # where the editor's own save dialogs open

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
        average = re.match(r"(.*)_avg(_[\d+]+|\d+)(_\d+)?$", name)
        if average:                                          # an average must not look like one of its measurements
            base = average.group(1)
            return "∑ " + (base if len(base) <= 28 else base[:27] + '…') + " (average)"
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
        files = [a for a in sys.argv[1:] if a.lower().endswith(('.nsz', '.csv'))]
        if files:
            app._load_files(files)

    root.mainloop()


if __name__ == '__main__':
    main()
