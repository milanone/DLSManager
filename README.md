# DLS-Horiba

Desktop viewer for `.nsz` files from the **Horiba SZ-100** dynamic light scattering (DLS)
instrument. It reads the binary files directly, without the instrument software, plots the
correlation function and the particle size distribution, compares several samples, and
exports everything to Excel.

> **Unofficial.** This project is not affiliated with or endorsed by Horiba. The `.nsz`
> format was worked out by inspecting files, and the extracted values were validated against
> the CSV exported by the instrument software on a single sample (see "Validation" below). Check the results
> against your own exports before relying on them, especially with other instrument models.

![Autocorrelogram tab](screenshot.png)
![Distribution tab](screenshot_distribution.png)
![DLS analysis tab](screenshot_dls.png)

## Features

- Opens `.nsz` files from a dialog, a whole folder, the command line or drag-and-drop
  (needs `tkinterdnd2`).
- Loads many samples and overlays the selected ones.
- **Autocorrelogram** tab: measured correlation function, Horiba fit and residuals. The file stores the intensity
  correlation g₂−1 (the square of the g₁ that the instrument software exports), and that is what is plotted.
- **Distribution** tab: size distribution by intensity or by number, log or linear x axis,
  optional cumulative curve.
- **DLS analysis** tab: the stored correlogram is analysed again in the program, and the result is shown next to the
  instrument's values. Cumulants (Z-average and PI) from a quadratic fit of ln g₁; an intensity-weighted size distribution
  either free (regularised non-negative least squares) or as one or two log-normal peaks with a free intercept of g₁
  and a BIC that says whether the second peak is supported. Needs `scipy`.
  The distribution can be shown by intensity, volume or number, converted with Mie theory (the light one particle scatters
  at the measurement angle, with the refractive index of the particle stored in the file or typed) or with the Rayleigh
  approximation (d⁶). The instrument's own values are shown only for intensity. See "Validation" for how far to trust it.
- **Parameters** tab: date, solvent, scattering angle, wavelength, temperature, viscosity,
  refractive index, peak and mean sizes, SD, D10/D50/D90, mode, median, Z-average, PDI and span.
- **Excel export**: one `.xlsx` per sample with sheets for the ACF (data, fit, residuals),
  the intensity and number distributions and the parameters.
- **Figures**: `File → Save Figure Image (PNG, PDF, SVG)...`, `Save Figure (pickle)...` and `Edit Figure...` work on the
  plot of the active tab (a copy restyled as an Origin-like figure with the shared
  [PlotStyleKit](https://github.com/milanone/PlotStyleKit), which is also used for the on-screen plots). PlotStyleKit is
  optional: the program looks for it next to itself (`../PlotStyleKit`), in `~/progetti/PlotStyleKit`, in the folder
  named by the environment variable `PLOTSTYLEKIT` and in the folder you chose the last time; if it finds none it says
  so at start and asks for the folder, with *Cancel* for going without it (it will not ask again; *File >
  PlotStyleKit...* changes that). Without it the plots keep a plain look with a dashed grid, *Save Figure* writes the
  figure as drawn and *Edit Figure* offers to locate PlotStyleKit.

## Run

```
pip install -r requirements.txt
pythonw nsz_viewer.pyw [file.nsz ...]
```

`pythonw` starts it without a console window (`python` keeps one).

## Tests

```
py -m unittest discover -s tests -v
```

The helpers and the application are tested on synthetic data. The parser is tested against a real measurement and the
CSV the instrument software exported for it, kept in `example data/` (gitignored); those tests are skipped if the
files are missing.

## Validation

On one sample (a single `.nsz` and the CSV of the same measurement) the viewer matches the software for: mean, SD, mode,
Z-average and PDI; the 84-bin size distribution (frequency and cumulative); and the correlation function and its fit
(the file holds g₁², the CSV g₁; the residuals follow the same pattern but are not equal). Not verified: other
instrument models or software versions, measurements with two peaks, and the date/time (the time is read from the
measurement ID and showed `12:22:00` for a measurement that the software dates `12:22:39`).

The DLS analysis was checked on 15 measurements of a polystyrene latex (certified 130 nm by TEM; 1:100 to 1:800 in 10 mM
NaCl, 173° and 90°): the Z-average of the cumulants is within 1% of the instrument's in 12 of them and within 4% in all,
the mode of the free distribution equals the instrument's at 173° (142 nm) and the one-peak model finds one narrow peak
(σ of ln d about 0.25) for which a second peak is not supported (BIC). At 90° the two-peak model does improve the BIC, but the
second peak (15-30%) sits at the width limit of the fit, which is not credible for a latex: read it as a tail, not a
population. The PI is noisy from file to file (it often differs from the instrument's by 0.03-0.08). On the very
polydisperse sample of `example data/` (PI 0.6) the Z-average is 4% above the instrument's, the free inversion piles up at
the upper size limit (3000 nm) and the peak fits hit their width limit: the methods are at the edge of their validity for
such a sample. The two-peak split is an ill-posed inversion: it is not stable between replicates and is only meaningful for sizes
that differ by about 3 times or more; the BIC difference is shown for that reason.

The correlator channels that the instrument used for the size distribution are read from the file (from channel 7 up to a
channel that differs between measurements) and applied to the inversions by default (*Instrument range*; the excluded points are
drawn hollow). On the files checked this changes the distributions very little, so it does not explain why the program's
free distribution of the 450 nm liposomes (mean 453 nm, mode 475 nm) differs from the instrument's (413 and 380 nm).

Volume and number weightings are conversions of the fitted intensity distribution, not measurements. Mie theory is the
default (homogeneous sphere, no absorption, vertical polarisation; the code reproduces Qext = Qsca and S(0) from an
independent computation with scipy's Bessel functions, but it was not compared with the instrument's own volume/number
tables, which are not read). Rayleigh (d⁶) is exact only well below 100 nm. Both divide by a quantity that falls steeply
with size, so the small-size side and any tail are amplified: for the free distribution a tail cut (default 5% of the
maximum) drops the few percent that the inversion leaves at the edge of the size grid, which otherwise becomes all of the
particles (on a 450 nm liposome sample the number mean goes from 10 nm without the cut to about 70-100 nm with it). With
Mie at 173° the per-particle intensity oscillates with size, so volume and number can change a lot with the particle
index. Liposomes are not homogeneous spheres of index 1.6, so for them the numbers are indicative only.

## Notes on the format

An `.nsz` file is an OLE compound file. The viewer reads these streams:

| Stream | Content |
|---|---|
| `Delay`, `OriginalACF` (inside one stream) | The 43 lag times (seconds) and the normalised correlation g₂−1, each array preceded by its byte length |
| `Object2`, `Object3` | Same correlation, but `Object3` lacks the first 6 lag times, so the program no longer uses it (kept as a fallback) |
| `Object10`, `Object7` | Fit and residuals; the header holds the number of initial channels excluded from the fit |
| `Object5`, `Object4`, `Object31` | Size bins (84), intensity distribution, cumulative distribution |
| `Base1`, `Base34`, `Base38` | Calculated results, measurement parameters, sample name |

Two details matter when interpreting the values: the diameter fields (`CalcMean`, `CalcPeakPos`,
`CalcPerDiameter`, `Cum_fMean`, ...) are stored as **radii**, so the viewer multiplies them by 2;
and the PDI is taken from `Cum_fPi`, because `CalcTotalPI` does not match the software's value.

## Data

This repository contains code only: `.gitignore` excludes `*.nsz`, `*.csv`, `*.xlsx` and the `example data/` folder,
so your measurements and exports stay on your machine. The screenshots show one real measurement.

## Requirements

Python 3, numpy, matplotlib, olefile, openpyxl; `scipy` (for the DLS analysis) and `tkinterdnd2` are optional.

## License

[MIT](LICENSE)
