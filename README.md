# DLS Manager

Desktop program for the `.nsz` files of the **Horiba SZ-100** dynamic light scattering (DLS) instrument
(formerly "NSZ Viewer"). It reads the binary files directly, without the instrument software, plots the
correlation function and the particle size distribution, compares several samples, analyses the stored
correlogram again (cumulants, free and log-normal size distributions by intensity, volume and number) and
exports everything to Excel.

> **Unofficial.** This project is not affiliated with or endorsed by Horiba. The `.nsz`
> format was worked out by inspecting files, and the extracted values were validated against
> the CSV exported by the instrument software on two measurements, one for each of the two file layouts below (see "Validation" below). Check the results
> against your own exports before relying on them, especially with other instrument models.

![Autocorrelogram tab](screenshot.png)
![Size distribution tab, source Instrument](screenshot_distribution.png)
![Size distribution tab, source Recomputed](screenshot_dls.png)

## Features

- Opens `.nsz` files from a dialog, a whole folder, the command line or drag-and-drop
  (needs `tkinterdnd2`).
- Loads many samples and overlays the selected ones.
- **Autocorrelogram** tab: measured correlation function, Horiba fit and residuals, as g₁ (what the instrument software
  shows and what the Size distribution tab plots) or as the g₂−1 that the file stores (g₁ squared).
- **Size distribution** tab, drawn the same way for both sources: the correlogram (g₁) with its fit, the size distribution
  and, optionally, the cumulative curve.
  - *Instrument*: the fit and the distribution that the instrument software stored in the file.
  - *Recomputed*: the stored correlogram is analysed again in the program, and the result is shown next to the
    instrument's values. Cumulants (Z-average and PI) from a quadratic fit of ln g₁; an intensity-weighted size
    distribution either free (regularised non-negative least squares) or as one or two log-normal peaks with a free
    intercept of g₁ and a BIC that says whether the second peak is supported. Needs `scipy`. See "Validation" for how far
    to trust it: the intensity distribution and the Z-average are checked against the instrument, volume, number and the
    two-peak fit are not (red cautions say so on the plot).

  Both sources can be shown by intensity, volume or number, on a log or linear axis. Volume and number are converted
  from the intensity distribution with Mie theory (the light one particle scatters at the measurement angle; solid sphere
  with the refractive index of the particle stored in the file or typed in *Particle n*; or a vesicle as a shell around a core
  of medium, with a number of lamellae, the bilayer thickness, the repeat distance and the lipid's index excess) or with the
  Rayleigh approximation (d⁶). The fields of the conversions not in use are greyed out, and Rayleigh shows a red warning when
  more than 5% of the intensity is above the size where d⁶ is off by 10% (60-90 nm for a solid sphere of index 1.6 in water). The instrument's own mean, mode and S.D.
  are shown only for intensity.
- **Cursor bar** on every plot: click or drag in a panel to put a vertical bar there; a box lists the value of each curve
  at the bar (the pan/zoom tools of the toolbar take precedence while active), right click removes it. The bar is left out
  of saved figures. In the distribution plot the vertical dotted line is the Z-average (the instrument's, or the one from
  the cumulants for the recomputed source) and the dashed lines of the cumulative plot are the instrument's D10, D50 and D90
  (all named in the legend).
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
pythonw DLSManager.pyw [file.nsz ...]
```

`pythonw` starts it without a console window (`python` keeps one).

`DLSManager.ico` is an icon for the program's files in a file manager (drawn by `examples/make_icon.py`, the same style as the
other programs' icons: a classic 16-64 px ICO).

## Tests

```
py -m unittest discover -s tests -v
```

The helpers and the application are tested on synthetic data. The parser is tested against a real measurement and the
CSV the instrument software exported for it, kept in `example data/` (gitignored); those tests are skipped if the
files are missing.

## Validation

The `.nsz` layout changed between software versions (the `Object` streams are numbered differently, the measurement
parameters moved from `Base34`/`Base38` to `Base35`/`Base39`), so the program finds the curves by their content. It was
checked against the CSV that the instrument software exported for the same measurement, on one file of each version (June
2026: a multilamellar liposome sample; 19 November 2026: `lipo 100 1a10 in PBS`). In both the program matches the software for:
mean, SD, mode, Z-average and PDI; the 84-bin size distribution (frequency and cumulative); and the correlation function and its
fit (the file holds g₁² in the older layout, the CSV g₁; the residuals follow the same pattern but are not equal). Not verified:
other instrument models, measurements with two peaks, and the date/time of the older file (the time is read from the
measurement ID and showed `12:22:00` for a measurement that the software dates `12:22:39`).

The DLS analysis was checked on 15 measurements of a polystyrene latex (certified 130 nm by TEM; 1:100 to 1:800 in 10 mM
NaCl, 173° and 90°): the Z-average of the cumulants is within 1% of the instrument's in 12 of them and within 4% in all,
the mode of the free distribution equals the instrument's at 173° (142 nm) and the one-peak model finds one narrow peak
(σ of ln d about 0.25) for which a second peak is not supported (BIC). At 90° the two-peak model does improve the BIC, but the
second peak (15-30%) sits at the width limit of the fit, which is not credible for a latex: read it as a tail, not a
population. The PI is noisy from file to file (it often differs from the instrument's by 0.03-0.08). On the very
polydisperse sample of June (PI 0.6, a multilamellar liposome preparation) the Z-average is 4% above the instrument's, the free inversion piles up at
the upper size limit (3000 nm) and the peak fits hit their width limit: the methods are at the edge of their validity for
such a sample. The two-peak split is an ill-posed inversion: it is not stable between replicates and is only meaningful for sizes
that differ by about 3 times or more; the BIC difference is shown for that reason.

The correlator channels that the instrument used for the size distribution are read from the file (from channel 7 up to a
channel that differs between measurements) and applied to the inversions by default (*Instrument range*; the excluded points are
drawn hollow). On the files checked this changes the distributions very little, so it does not explain why the program's
free distribution of the 450 nm liposomes (mean 453 nm, mode 475 nm) differs from the instrument's (413 and 380 nm).

The reason is that the instrument's distribution is not an inversion of its own fit. On the 29 intensity distributions
checked (latex and liposomes, PI from 0.014 to 0.38) it is always a peak of the same width, 10-11 bins with a σ of ln d of
0.22-0.28, and only its position changes. Passed through the scattering kernel, it does not reproduce the correlogram that
the instrument fits to the data (rms of g₂−1 0.004 for the latex and 0.03 for the 450 nm liposomes, against 0.0003 for the
instrument's own fit), whereas a free log-normal peak fitted to the data reaches 0.0002-0.001 with σ of 0.38-0.77. The
instrument's width therefore says nothing about the sample, while the program's follows the data (and is limited by its
resolution). Also, the diameters in the instrument's table are bin upper edges: its mean is exactly the table's mean
divided by √1.1298 (the bin ratio) in all 29 files, so its stated mode and mean sit about 6% below the plotted peak.

The Mie conversion can average the light one particle scatters over a detector aperture (an entry in the tab, 0° = the single
angle by default; the SZ-100's own value is not known), because at a single angle the interference minima of Mie theory (for a latex at 90°, one
near 350 nm) turn the small tail of a free distribution into a spurious peak in volume and number. It reduces that peak but does not
remove it: on the 130 nm latex the volume distribution still has a smaller peak near 350 nm at ±10° and ±20°, and it only vanishes at
±30° or more, wider than a real detector, so it is a smoothing and not a correction (hence the default of 0°). On synthetic data built from the lot certificate (number distribution 130 ± 14 nm) the one-peak model recovers the number
and volume means (130 and 134 nm), while the free distribution widens the peak and gives a number mean 27% low: volume and number
depend on the width of the distribution, which DLS resolves worst.

The instrument's software computes its size distribution with a *Form of Distribution* (Edit > Recalculate): Monodisperse or
Polydisperse, and Broad, Standard, Narrow or a manual *resolution* from -1 to 1 (presets in its manual: Monodisperse -0.5, 0.3,
0.6; Polydisperse -0.3, 0.1, 0.3). Every measurement checked was in the Monodisperse Standard form (resolution 0.3, stored in the
file as `Dist_Resolution`). On the size tables that the software recalculated from the same correlogram, the resolution fixes
the WIDTH of the peak it shows, whatever the sample (σ of ln d 0.43, 0.24 and 0.07 for Broad, Standard and Narrow; the fit of
the correlogram is equally good in all of them), and Polydisperse allows several peaks: the 450 nm liposomes give two,
157 nm (39%) and 611 nm (61%), which the program's two-peak model reproduces (172 nm 41%, 588 nm 59%). A recalculation changes
only the distribution (and its fit curve): Z-average, PI and the correlogram stay the same.

The *Width (instrument form)* choice (free distribution only) gives the computed distribution the width of one of these forms
around its centre, in intensity and before the conversion to volume and number: *As computed* leaves it as the inversion gives
it (default), *Broad*, *Standard* and *Narrow* set σ to 0.43, 0.24 and 0.07, and *Manual* takes a resolution (the map from
resolution to σ goes through the four points measured on the recalculated tables, so it is an interpolation). With Standard on
the latex and the 100 nm liposomes the mean, mode and S.D. agree with the instrument's table within about 3% in intensity,
volume and number (Mie conversion); on the 450 nm liposomes the instrument's peak sits 10-15% higher than the program's, and
how it places the peak is not known. The regularised inversion widens a narrow sample, and volume and number depend on the
width through roughly exp(−kσ²): on the latex of the lot certificate (number mean 130 nm) a narrow form gives number means of
127-136 nm and a single volume peak, against 91-106 nm and a spurious second volume peak as computed. These forms are
assumptions about the sample, not measurements: a truly broad one (extruded liposomes with a PI of 0.4) is narrowed wrongly.
The choice is shown in the text under the plot, and the narrowed curve no longer follows the correlogram fit drawn beside it.

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
| fit and residual streams (`Object10`, `Object7` or `Object11`, `Object8`) | Fit and residuals of the channels the instrument used for the distribution (from channel 7) |
| size bins, intensity and cumulative distribution (`Object5`, `Object4`, `Object31` or `Object6`, `Object5`, `Object32`) | 84 bins each |
| `Base1`, `Base34`/`Base35`, `Base38`/`Base39` | Calculated results, measurement parameters, sample name |

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
