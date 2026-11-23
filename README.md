# Microplastic Data Processor

A Shiny app that turns raw per-particle exports from **FTIR**, **Raman**, and
**LDIR** microspectroscopy instruments into formatted, analysis-ready Excel
workbooks.

## What it does

Each instrument's particle-analysis software exports a table of detected
particles (material identification, size, area, ...), but in a different
shape and with inconsistent material naming. This app:

1. **Auto-detects the instrument** from the uploaded file's column
   signature (`Eccentricity` → LDIR, `RamanSignal` → Raman, an `Area on map`
   column → FTIR; FTIR is further split into Spotlight/Lumos sub-types from
   the filename).
2. **Standardizes material names** — messy free-text identifications (e.g.
   "BIAXIALLY-ORIENTED EVOH FILM", "Polypro", "polyethylene terephthalate")
   are mapped to a fixed set of polymer-group abbreviations (PE, PP, PS, PA,
   PVC, PET, PC, ABS, ...), with unmapped names reported separately rather
   than silently dropped.
3. **Bins particles by size** (Feret / CE diameter) into four size classes.
4. **Builds a multi-sheet Excel workbook** per input file (or a combined
   workbook across every file in a folder): a long-format particle table,
   per-file counts by polymer type and size class, mass totals (FTIR),
   folder-wide totals and per-file means, a material name-mapping audit
   sheet, and — for Raman — an "unmatched material" sheet plus a processing
   log.

The interface handles multi-file batches via drag-and-drop or a folder
picker, shows a live processing log, and includes a guided in-app tour.

## For users

You only need [R](https://cran.r-project.org/bin/windows/base/) installed
(ask IT if it isn't), plus an internet connection.

**One-time setup:** download
[`Start Microplastic Processor.bat`](https://github.com/KristofM854/Excel-processor-FTIR-Raman/raw/main/Start%20Microplastic%20Processor.bat)
and copy it to your desktop.

**Every day:** double-click it. A black window opens and the app starts in
your browser. Keep the black window open while you work, and close it when
you're done.

**Updates happen automatically every time you start it.** The launcher
always downloads the latest version, so there is nothing to install or
replace. The version you are running (date and code) is shown next to the
app title.

The first start can take a few minutes because R installs the packages the
app needs. Later starts are quick. If R is missing or GitHub can't be
reached, the window says so.

## Running it (developers)

```r
# Installs any missing packages automatically on first run
shiny::runApp("app.R")

# Or exactly what the desktop launcher does (latest version on main):
shiny::runGitHub("excel-processor-ftir-raman", "KristofM854", ref = "main", launch.browser = TRUE)
```

Whatever is merged to `main` reaches every user the next time they start the
launcher, so only merge versions that are ready to use.

Requires `shiny`, `shinyFiles`, `dplyr`, `tidyr`, `openxlsx`, `stringr`,
`readr`, `janitor`, `readxl`, `rintrojs`. `app.R` installs whichever of
these are missing when it starts.

### Input files

| Instrument | Format | Required columns |
|---|---|---|
| FTIR (Spotlight/Lumos) | `.csv` | `Identifier`, `Group`, `Feret min [µm]`, `Mass [ng]`, `Area on map [µm²]` |
| Raman | `.csv` | a material/polymer column, an HQI (quality index) column, a Feret-max column |
| LDIR | `.xlsx` (sheet 2) | `Eccentricity`, `Is Valid`, `Match Type`, `Id`, `Area (µm²)` |

CSV encoding (UTF-8, UTF-16LE/BE, Windows-1252, Latin-1) and delimiter
(comma or semicolon) are auto-detected per file.

## Testing

`tests/test-standardize_group_code.R` is a standalone check of the material
name-mapping rules — source it from the project root:

```r
source("tests/test-standardize_group_code.R")
```

## License

MIT — see [LICENSE](LICENSE).
