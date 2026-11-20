# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.
General rules for all of Francesco's projects (stack, code conventions, repository rules, git/GitHub, the
Drive mirror, how he works) are in `..\CLAUDE.md`; this file has the project-specific details.

## Purpose

Tkinter viewer for the `.nsz` files of the **Horiba SZ-100** dynamic light scattering (DLS) instrument: DLS analysis of the correlogram, correlation
function, size distribution, parameters table, multi-sample overlay, Excel export, figure save/edit. The format was
reverse-engineered, so the parser is checked against the CSV exported by the instrument software.

## Running / tests

```bash
pythonw nsz_viewer.pyw [file.nsz ...]
py -m unittest discover -s tests -v
```

Sample data: `example data/lipo_bup_MLV_0minson_scattering_0729.nsz` and the software CSV of the same measurement
`..._0731.csv` (gitignored: real data, never commit). Tk tests need a screen; tests that need the sample are skipped
without it. The two `screenshot*.png` are the real program with that sample (`ImageGrab` on the window, topmost).

## Architecture

Single file `nsz_viewer.pyw`: parser functions (`f32`, `parse_props`, `load_nsz` / `_read_nsz`, `dump_nsz` for
debugging) and the class `NSZViewer` (sidebar with sample list, a `ttk.Notebook` with Autocorrelogram / Distribution /
Parameters tabs, two matplotlib `Figure`s `fig_acf` and `fig_size`). `self.samples = {path: data dict}`, `self.selected`
= paths shown. `_plot_acf` / `_plot_size` redraw on every control change, `_update_params_table` fills the table (up to 5
samples), `_write_xlsx` writes one workbook per sample, `_to_number_dist` converts intensity to number distribution
(Rayleigh, N proportional to I/d^6).

## The .nsz format (OLE compound file, read with `olefile`)

| Stream | Content and how it is read |
|---|---|
| `Object3` | lag times: 16 header floats, then tau (µs); trailing repeated baseline channels are cut where tau stops increasing (43 -> 37 channels in the sample) |
| `Object2` | correlation (same numbers as `OriginalACF`): 3 header floats, then one value per channel. These are **g1 squared** (g2-1): the software's CSV g1 equals `sqrt(value)` to 5e-7, CSV row k = channel 7 + k |
| `Object10` | fit: header `[1][n_fit][fit_offset]`, then values; `fit_offset` (6) = initial channels excluded from the fit; also g1 squared, CSV fit row k = `sqrt(fit[k+1])` |
| `Object7` | residuals (same offset as the fit); correlated (0.83) with the CSV residual of g1, not equal |
| `Object5` | diameters: 17 header floats, then 84 bins |
| `Object4` | intensity distribution: 3 header floats, then 84 values |
| `Object31` | cumulative distribution: 4 header floats, then 84 values |
| `Base1`, `Base34`/`Base35`, `Base38`/`Base39` | calculated results, measurement parameters, sample name; `parse_props` reads `name\0` + float32 pairs. **The numbers depend on the software version** (files of 2026-11-19 use Base35/Base39): `_read_nsz` picks the stream by content (`MeasAngle`, `Sample_Name`). Before this fix (2026-11-20) those files showed no angle, temperature, viscosity or solvent |

Details that matter: diameter fields (`CalcMean`, `CalcPeakPos`, `CalcPerDiameter`, `Cum_fMean`, ...) are stored as
**radii** (x2); the PDI is `Cum_fPi` because `CalcTotalPI` does not match the software; `MeasID` holds date and time.

**Bug fixed while writing the tests (2026-10-08):** the three distribution streams were read with the wrong header
lengths (16, 2, 2 instead of 17, 3, 4), so the diameter and intensity arrays were shifted by one bin (a spurious 0.30 nm
first bin, the last bin 8510 nm dropped) and the cumulative curve was shifted by one bin against the diameters. The
header lengths above reproduce all three CSV columns exactly (to the CSV rounding).

## DLS analysis (tab `DLS analysis`, ported from LabSpectrumManager)

Module-level functions after the parser: `dls_q2`, `_dls_gamma`, `dls_cumulants` (quadratic fit of ln g1, g1 > 0.3), `dls_invert`
(regularised NNLS, 10-3000 nm, alpha 1, g1 > 0.1), `dls_fit_peaks` (1 or 2 log-normal peaks, free intercept of g1, delays below
1 µs left out because the stored ACF drops ~1% after its first point and a fit reads that as a 3-5% peak at 3-16 nm, sigma >= 0.05,
BIC), `dls_analyse`, `dls_record(d)` (delays in seconds and the conditions read from the file: angle, solvent n, viscosity,
temperature; None if the delays are not aligned). Wavelength is an entry (532, also `Lambda_nm` of the file). `NSZViewer._plot_dls`
draws g1 with the fit and the intensity distribution for all selected samples and `_dls_summary` writes the comparison with the
instrument (Z-average, PI, mean, mode, S.D.). It runs only while the tab is shown (`_refresh_dls`, `<<NotebookTabChanged>>`) and caches
results in `_dls_cache`; changing the model resets g1 min (0.1 regularised, 0.05 peaks). `ToolTip` gives the help texts (`DLS_TIPS`).
The figure is saved/edited like the others (`fig_dls`, preset `double`). Tests: `tests/test_dls.py` (synthetic correlograms of
known sizes, the real sample, the tab).

Weighting selector (`dls_reweight`, `dls_reweight_peaks`, `dls_stats`, `dls_floor`, `DLS_WEIGHTS`, `DLS_CONVERSIONS`): number = p / I1,
volume = p d^3 / I1 on the log grid, with I1 = light one particle scatters, from `mie_intensity` (Bohren-Huffman, |S1|^2, vertical
polarisation, homogeneous non-absorbing sphere, index `MeasSampleRef` of the file = `params['n_particle']`, 1.6 default in the
instrument, editable in the tab) or d^6 (Rayleigh, as `_to_number_dist`). Log-normal peaks: analytic for Rayleigh (median moves by
exp(-k sigma^2), share by exp(-k mu + k^2 sigma^2 / 2)), rebuilt on the grid for Mie (`dls_reweight_peaks(..., d, per_particle)`).
`TestMie` checks d^6 at small sizes and Qext = Qsca and S(0) against scipy's spherical Bessel functions (independent code), not
against the instrument. The free inversion leaves ~2% intensity on the 10 nm grid edge: unconverted, it becomes ALL the
particles (number mean 10 nm on the 450 nm liposomes), so `dls_floor` drops bins below the 'Tail cut' (5% of the maximum, an entry in the
tab, 0 keeps everything) before volume/number; this is applied to the free model only. Instrument values are shown only for
intensity (its number/volume tables were never located). Mie at 173 degrees oscillates with size (latex volume mode 252 nm
against 142 in intensity), and liposomes are shells, not solid spheres: numbers for them are indicative. Controls are on two rows
(they did not fit the window on one). Not compared: any instrument volume/number output.

Verified on 15 polystyrene latex measurements (the `20261119 DLS lattice e liposomi` folder of LabSpectrumManager, 130 nm by TEM;
`tests/test_dls.py::TestLatex` uses the three at 173 degrees and skips without them): Z-average -4.0..+0.3% of the instrument's
(within 1% in 12), regularised mode 142 nm = the instrument's at 173 degrees, one peak median 140 nm with sigma ~0.25 and no
support for a second peak at 173 degrees (BIC difference -8..-9); at 90 degrees the 2-peak fit wins on BIC but its second peak
(14-31%) is at the sigma bound 1.2, so it is a tail, not a population. PI is within ~0.03-0.08 of the instrument's, either way.
Also verified in LabSpectrumManager on four files; the two-peak split is ill-posed and only sensible for sizes ~3x apart.
Verified on the `example data` sample (PI 0.6): Z-average +4.2%, PI 0.44 against 0.60; the free
inversion has its mode at the 3000 nm limit and the peak fits reach sigma 1.2 (bound), so for this aggregated sample the sizes of
the distribution are not reliable beyond the Z-average. Not done: the DLS results are not in the Excel export.

## Verified / not verified

- Verified on **one** measurement against the software CSV: Mean, SD, Mode, Z-average, PDI; size table (diameter,
  frequency, undersize); correlation and fit (as g1 squared). See `tests/test_nsz_viewer.py::TestRealSample`.
- The correlation is plotted as stored (g2-1, label `g2(tau) - 1`); the software shows g1. Offer to switch to g1
  (`sqrt`) if Francesco wants the same curve as the software.
- Not verified: other SZ-100 software versions or instruments, two-peak distributions (`Peak2`/`Mean2` exist in the
  parser but were never seen non-empty), and the time of day: `MeasID` gave `12:22:00` where the software prints
  `12:22:39`.
- `Temp_C` is 24.95 where the software prints 25.0 (rounding of a displayed value).

## PlotStyleKit

Found by the common bootstrap block near the top of `nsz_viewer.pyw` (the same text as in PlotStyleKit's README: change it
there, not here): `psk` is its integration module `plotstylekit.py` or None, `PSK_PROBLEM` says why. If it is not found and
the user has not declined, `_offer_plotstylekit()` asks for the folder 400 ms after the window opens; File > PlotStyleKit...
(`plotstylekit_dialog()`) changes it later and replots (`_refresh_all`, because the sample colours change too).
`sample_colors()` uses the Origin colour cycle (black first) with it, else the fallback `COLORS`. `_style_figure(fig)` applies
`psk.style(ax, None, set_size=False)` to every axes, enlarges the fonts for the pane and runs one `tight_layout` (no grid with
the style; the dashed grid only without PlotStyleKit). File menu: `Save Figure Image`, `Save Figure (pickle)`, `Edit Figure...`
act on the active plot tab (`_current_figure()`) through `psk.save_*_dialog` / `psk.edit_dialog` (`preset` = `single` 4:3 for
one panel, `double` 16:9 for two: `_preset_for`; `unpin=False`: the legend stays where it is). Fallbacks without it: the figure
is saved as drawn (`_copy_figure()` is a plain pickle copy) and 'Edit Figure' offers to locate PlotStyleKit
(`_need_plotstylekit`). Style changes go in PlotStyleKit, never here. (This project exposed that `applica_stile_origin` set
`AutoMinorLocator` on log axes, which matplotlib warns about; PlotStyleKit now skips non-linear axes.)
`tests/test_plotstylekit.py` drives all of this with answers in place of the dialogs.

## Editing conventions (project)
- Follow `..\CLAUDE.md`: everything in English, surgical edits. The UI was in Italian and was translated in one pass
  (menus, labels, plot titles, Excel sheet names: `Distribuzione intensità` is now `Intensity distribution`, etc.).
- Never commit anything from `example data/`.
