# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.
General rules for all of Francesco's projects (stack, code conventions, repository rules, git/GitHub, the
Drive mirror, how he works) are in `..\CLAUDE.md`; this file has the project-specific details.

## Purpose

Tkinter program (**DLS Manager**, formerly NSZ Viewer; renamed 2026-11-20 together with the file `DLSManager.pyw` and the class `DLSManager`; the repository and folder were renamed from DLS-Horiba / DLS Horiba to DLSManager the same day) for the `.nsz` files of the **Horiba SZ-100** dynamic light scattering (DLS) instrument: DLS analysis of the correlogram, correlation
function, size distribution, parameters table, multi-sample overlay, Excel export, figure save/edit. The format was
reverse-engineered, so the parser is checked against the CSV exported by the instrument software.

## Running / tests

```bash
pythonw DLSManager.pyw [file.nsz ...]
py -m unittest discover -s tests -v
```

Sample data: `example data/` holds `lipo 100 1a10 in PBS_0843.nsz` and the software CSV of the same measurement
(`..._0843.csv`; gitignored: real data, never commit). On 2026-11-20 the folder had a CSV named `_0845` next to the `_0845` .nsz, but
the CSV was of measurement 0843 (its header, its time 18:03:04 and all its values match the 0843 file, not 0845): Francesco
replaced the pair with 0843 + its CSV. `find_sample()` in `tests/test_nsz_viewer.py` reads the .nsz named in the CSV (next to it, else in
the `spettri di esempio` folders of LabSpectrumManager). The earlier sample (`lipo_bup_MLV_0minson_scattering_0729.nsz` + `_0731.csv`, June
layout, PI 0.6) is gone from the folder; its results are described below. Tk tests need a screen; tests that need the sample are skipped
without it. The three `screenshot*.png` are the real program with the 0843 file (window grabbed with `CopyFromScreen`, cropped).

## Architecture

Single file `DLSManager.pyw`: parser functions (`f32`, `parse_props`, `load_nsz` / `_read_nsz`, `dump_nsz` for
debugging) and the class `DLSManager` (sidebar with sample list, a `ttk.Notebook` with Autocorrelogram /
Size distribution / Parameters tabs, matplotlib `Figure`s `fig_acf`, `fig_size` and `fig_dls`). `self.samples = {path: data dict}`, `self.selected`
= paths shown. `_plot_acf` / `_plot_size` redraw on every control change, `_update_params_table` fills the table (up to 5
samples), `_write_xlsx` writes one workbook per sample, `_to_number_dist` converts intensity to number distribution
(Rayleigh, N proportional to I/d^6).

## The .nsz format (OLE compound file, read with `olefile`)

| Stream | Content and how it is read |
|---|---|
| *layouts* | **The `Object` numbers depend on the software version.** The 2026-06 file: sizes `Object5`, intensity `Object4`, cumulative `Object31`, fit `Object10`, residuals `Object7`. The 2026-11-19 files: `Object6`, `Object5`, `Object32`, `Object11`, `Object8` (`Object4` is a 114 kB blob). `_nsz_curves()` finds them by content (101 floats ending near 8510 = diameters; 88 floats summing to 100 = intensity, rising to 100 = cumulative; the stream of 3 + (`Dist_ChannelTo` - `Dist_ChannelFrom` + 1) floats closest to the ACF = fit g2-1, the small-valued one most correlated with ACF - fit = residuals) and the old names are only a fallback. Before this (fixed 2026-11-20) the newer files gave 12 size bins and a wrong fit. If neither layout is recognised the fit is left empty (`_plausible_fit`): Francesco once saw a spike at 5 us because his running copy predated the fix, and the old reader returned a 16-float header as the fit; restart the program after updating it. Verified against the CSV of `_0843` (new layout): Z, PI, mean, mode, the three tables, correlation and fit (5e-7) and the residual pattern (0.88 at offset 1); and again the June file |
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

## Size distribution tab (instrument result and DLS analysis, ported from LabSpectrumManager)

One tab, `tab_size` ("Size distribution"; merged on 2026-11-20 from the former 'Distribution' and 'DLS analysis' tabs, which showed two
intensity distributions and two different number conversions). Built by `_build_size_tab`: the shared row has `size_source`
(`SIZE_SOURCES` = Instrument / Recomputed), `size_xscale`, `dls_weight` (intensity / volume / number) and the conversion
(`dls_conv`, particle n, shell thickness and index); `box_inst` holds the cumulative checkbox and `fig_size` (the file's own
84-bin distribution, `_plot_size`), `box_rec` the model controls, `fig_dls` and the text (`_plot_dls`). `_on_source` swaps the boxes,
`_plot_dist` redraws the visible one (`_refresh_dls` plots the recomputed one only while the tab is shown), `_current_figure`
returns `fig_size` or `fig_dls` by source. In the Instrument source volume/number use the same conversion as the Recomputed one (`dls_reweight`
with the per-particle intensity on the file's diameters; Rayleigh if the file lacks angle/index); `_to_number_dist` remains for the Excel export.
Cursor bar: `PlotCursor` (one per figure: `cur_acf`, `cur_size`, `cur_dls`): left click/drag puts a red vertical line in the clicked
panel, with a box listing each labelled curve's interpolated value (curves with a label starting with `_` are named by the attribute
`_cursor_name`, set on the unlabelled fit lines); ignored while a toolbar tool is active (`toolbar.mode`); right click removes it. The
plot methods rebuild the figure with `fig.clear()`, so each calls `cur.attach()` before `canvas.draw()`; the decorator `_no_cursor` takes
the bar out of the figure around Save Figure Image / Save Figure (pickle) / Edit Figure (they copy the figure). Vertical lines in the plots:
D50 (dotted, instrument distribution), D10/D50/D90 (dashed, cumulative), Z-average (dotted, recomputed distribution), each with a
proxy entry in the legend. Tests: `TestApplication.test_cursor_*` send matplotlib mouse events.

The Recomputed part, in detail:

Module-level functions after the parser: `dls_q2`, `_dls_gamma`, `dls_cumulants` (quadratic fit of ln g1, g1 > 0.3), `dls_invert`
(regularised NNLS, 10-3000 nm, alpha 1, g1 > 0.1), `dls_fit_peaks` (1 or 2 log-normal peaks, free intercept of g1, delays below
1 µs left out because the stored ACF drops ~1% after its first point and a fit reads that as a 3-5% peak at 3-16 nm, sigma >= 0.05,
BIC), `dls_analyse`, `dls_record(d)` (delays in seconds and the conditions read from the file: angle, solvent n, viscosity,
temperature; None if the delays are not aligned). Wavelength is an entry (532, also `Lambda_nm` of the file). `DLSManager._plot_dls`
draws g1 with the fit and the intensity distribution for all selected samples and `_dls_summary` writes the comparison with the
instrument (Z-average, PI, mean, mode, S.D.). It runs only while the tab is shown (`_refresh_dls`, `<<NotebookTabChanged>>`) and caches
results in `_dls_cache`; changing the model resets g1 min (0.1 regularised, 0.05 peaks). `ToolTip` gives the help texts (`DLS_TIPS`).
The figure is saved/edited like the others (`fig_dls`, preset `double`). Tests: `tests/test_dls.py` (synthetic correlograms of
known sizes, the real sample, the tab).

Instrument range: the files store the correlator channels the instrument used for the size distribution as int32 `Dist_ChannelFrom` (7 in
every file seen, i.e. 1 us) and `Dist_ChannelTo` (26-40, different per measurement; for the 450 nm liposomes 33 = `len(fit) + fit_offset`
of Object10; for the 2026-11-19 files Object10 does not follow it). `params['Range_from'/'Range_to']` (1-based, inclusive),
`dls_record()['range']` (0-based [first, end)), `dls_analyse(..., use_range=True)` passes tau_min/tau_max to the inversions (cumulants
unrestricted); the tab's 'Instrument range' checkbox (on by default) draws the excluded points hollow. Effect on the results is small
(the g1 minimum already cuts the tail): 450 nm liposomes, mean/mode 453/475 with and without; it does not explain the gap to the
instrument's 413/380. `ChannelsCumuFrom/To` (1-41) and `FlAutoChannels` (2) are also in the file, not used.

Lamellae (2026-11-20): `_shell_params()` = (lamellae n, bilayer 4.5 nm, repeat distance 7 nm, lipid index excess 0.12; invalid -> these defaults), `_shell()` = the one
equivalent shell: thickness n*rep-(rep-t), index excess dn*n*t/thickness (LabSpectrumManager's `vesicle_amplitude` model; n = 1 is the
bilayer itself). 'Particle n' belongs to the 'Mie' conversion only; the fields of the other conversions are disabled (`_conv_fields`,
`_on_conv`). `_rayleigh_limit()` finds the diameter where Mie/d^6 drops below 0.9 of its small-size value (87 nm for the 450 nm
liposomes at 90 degrees, index 1.6; from the table: 10% off at ~60 nm at 90 degrees and ~50 nm at 173 degrees) and `_rayleigh_warning()` shows a red
note (`ray_warn`) when Rayleigh is used for volume/number and more than 5% of the intensity lies above it. Real-window checks
use the process's own window handle: a Tk combobox list is a popup that replaces `MainWindowHandle` while it is open.

Mie shell conversion (`mie_shell_intensity`, `DLS_CONVERSIONS` 'Mie shell'): a vesicle as a coated sphere with a core of the medium, shell thickness (default 4.5 nm) and index excess (0.12), the defaults and the coefficients of LabSpectrumManager's vesicle model; the tab has the two entries on a third row. Checked: with a shell thicker than the radius it equals the solid-sphere Mie to 1e-6, and a thin shell scatters ~d^4 at small sizes (instead of d^6). Volume stays the outer-sphere volume d^3 (not the lipid volume). The shell parameters are order-of-magnitude assumptions (a multilamellar vesicle needs a larger thickness), and no instrument value was compared.

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

The 'lipo 450 nm' samples (files `0826`-`0828`) are liposomes made by hydration and extruded three times through a 450 nm filter
(the one used for sterilisation), most probably multilamellar: Z-average about 300 nm (PI 0.38) while the mean of the distribution is
near the nominal 450 nm (ours 453, the instrument's 413). The Z-average of a broad distribution is the intensity-weighted harmonic mean and
sits low, so the two are not in conflict. For these the Mie shell thickness of 4.5 nm (one bilayer) does not apply: LabSpectrumManager estimated
about 7 lamellae for the same kind of sample.

Verified on 15 polystyrene latex measurements (the `20261119 DLS lattice e liposomi` folder of LabSpectrumManager, 130 nm by TEM;
`tests/test_dls.py::TestLatex` uses the three at 173 degrees and skips without them): Z-average -4.0..+0.3% of the instrument's
(within 1% in 12), regularised mode 142 nm = the instrument's at 173 degrees, one peak median 140 nm with sigma ~0.25 and no
support for a second peak at 173 degrees (BIC difference -8..-9); at 90 degrees the 2-peak fit wins on BIC but its second peak
(14-31%) is at the sigma bound 1.2, so it is a tail, not a population. PI is within ~0.03-0.08 of the instrument's, either way.
Also verified in LabSpectrumManager on four files; the two-peak split is ill-posed and only sensible for sizes ~3x apart.
Verified on the June sample (`lipo_bup_MLV..._0729`, PI 0.6): Z-average +4.2%, PI 0.44 against 0.60; the free
inversion has its mode at the 3000 nm limit and the peak fits reach sigma 1.2 (bound), so for this aggregated sample the sizes of
the distribution are not reliable beyond the Z-average. Not done: the DLS results are not in the Excel export.

## Verified / not verified

- Verified on **two** measurements (one per software version) against the software CSV: Mean, SD, Mode, Z-average, PDI; size table (diameter,
  frequency, undersize); correlation and fit (as g1 squared). See `tests/test_nsz_viewer.py::TestRealSample`.
- The correlation is plotted as stored (g2-1, label `g2(tau) - 1`); the software shows g1. Offer to switch to g1
  (`sqrt`) if Francesco wants the same curve as the software.
- Not verified: other SZ-100 software versions or instruments, two-peak distributions (`Peak2`/`Mean2` exist in the
  parser but were never seen non-empty), and the time of day: `MeasID` gave `12:22:00` where the software prints
  `12:22:39`.
- `Temp_C` is 24.95 where the software prints 25.0 (rounding of a displayed value).

## PlotStyleKit

Found by the common bootstrap block near the top of `DLSManager.pyw` (the same text as in PlotStyleKit's README: change it
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
