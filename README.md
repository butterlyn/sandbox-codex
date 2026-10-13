# 2026 WEM ESOO · Interactive data register

A single, self-contained HTML explorer recreating the figures in AEMO’s **2026 Wholesale Electricity Market Electricity Statement of Opportunities** data register.

**[Download index.html](https://raw.githubusercontent.com/butterlyn/sandbox-codex/refs/heads/feat/wem-esoo-2026-interactive-charts/index.html)** and open the saved file in a current Chrome, Edge, Firefox or Safari browser. The plotting library, chart values, and reference images are embedded. There are no runtime dependencies, external fonts, API requests, or analytics. Source-publication links require an internet connection.

All implementation is on **`feat/wem-esoo-2026-interactive-charts`**. `main` contains only the empty repository initialization commit.

## Coverage

| Content | Count | Treatment |
| --- | ---: | --- |
| Valid native Excel chart objects | 66 | Interactive recreation |
| Additional numerical figures stored as images | 14 | Interactive recreation from their worksheet data |
| Diagrams, maps and duplicate overview illustrations | 15 | Original images with zoom controls |
| Broken legacy Excel chart objects | 2 | Explicit source-issue entries; placeholders are not plotted as forecasts |
| **Total** | **97** | Every figure worksheet covered |

The register contains **68 native chart objects**. Some figures have multiple panels and some sheets contain duplicate or legacy objects, so chart-object counts differ from report figure numbers. The demand-traces XLSB has no embedded charts; it is downloaded and retained alongside the report and appendices.

The explorer supports chapter and text filters, links to individual figures, hover values, pan/zoom, series toggles and isolation, data tables, full CSV exports, PNG exports, and source references. Large availability curves retain **560,640 observations** in total. CSVs include plotted coordinates and original values/context where a display-unit conversion or derived series requires it.

## Sources

The supplied, revision-pinned AEMO files are retained in [`data/sources`](data/sources). Their original URLs, sizes, download timestamps and SHA-256 digests are recorded in [`sources.json`](data/sources/sources.json). The 2026 edition was confirmed on AEMO’s [WEM ESOO publication page](https://www.aemo.com.au/energy-systems/electricity/wholesale-electricity-market-wem/wem-forecasting-and-planning/wem-electricity-statement-of-opportunities-wem-esoo).

Source material © AEMO 2026. This is an independent recreation, not an AEMO product. AEMO’s source-document notices continue to apply. Plotly.js 3.1.0 is vendored under its [MIT license](vendor/plotly-LICENSE).

## Source irregularities and reconstruction choices

- Native chart objects **21 and 23**, on `Ch 3_F.24` and `Ch 3_F.25`, contain only broken external `#REF!` references and placeholder values of 1. They remain in the inventory as unavailable. The genuine charts on those sheets, objects 22 and 24, are interactive.
- Objects **2 and 19** contain three invalid placeholder series apiece. Those values are retained for audit and omitted from the forecast graphic. A broken category reference in object 1 is repaired from an unambiguous sibling series without changing its values.
- Saved chart caches are used to reproduce native charts. Hidden rows/columns follow Excel’s `plotVisOnly` setting. Missing values remain missing. Native display units, secondary axes, cumulative stacked lines, scenario groupings, hatch patterns, and seasonal gaps are preserved through documented transformations.
- Image-backed charts are reconstructed from their worksheet data. Appendix Figure 25 contains worksheet dates beyond its image’s stated period and some cause classifications differ from the image; all worksheet observations and classifications are retained and the discrepancy is disclosed. A percentage stored as text in a network simulation cell is normalized to the same percentage-point units as neighbouring values.
- Boxplots use all provided simulations with linear quartiles and 1.5-IQR whiskers. Original figure images are available in the **Source & notes** view for comparison. Layouts and font spacing are adapted for an interactive browser; they are not pixel-for-pixel Excel screenshots.
- Chart 66 preserves the workbook’s unusual cumulative stacked-line construction, with the underlying source values available for inspection. Derived regression lines and reserve-utilisation layers are explicitly identified in their figure notes.

Chart-specific formulas, repairs, omissions and assumptions are shown within the HTML and in the generated JSON specs.

## Rebuild

Python 3.10+ is sufficient for generation:

```sh
python -m pip install -r requirements.txt
python scripts/rebuild.py
```

The rebuild verifies/downloads the four originals, inventories the workbook, runs nine independent chart generators with up to six workers, validates coverage and source values, and combines everything into `index.html`. It does not require a Node build system. The deterministic gzip payload keeps the HTML around 18 MB while retaining the underlying observations.

To assemble existing chart specs after an HTML/CSS/JavaScript change:

```sh
python scripts/build.py
```

To serve the file locally:

```sh
python -m http.server 8000
```

Then open `http://localhost:8000`.

## Validation and parallel work

Six subagents worked concurrently: five native-chart groups (1–14, 15–28, 29–41, 42–54 and 55–68), plus an image-backed figure audit/reconstruction. Finished groups then handled the remaining supplemental network and appendix charts. The parent task integrated the common interface and single-file build.

```sh
python scripts/validate.py
node --check src/app.js
python -m pip install playwright
python scripts/browser_check.py
```

The browser check uses an installed Chromium browser; see its command-line options for configuration. This managed environment blocks `file://` navigation, so browser testing uses a local static server and rejects external network requests. Direct file opening could not be verified under that browser policy. The generated file contains all runtime assets and data, and needs a browser supporting `DecompressionStream`.

Validation outputs are in [`data/analysis`](data/analysis), including workbook coverage, native cache-versus-cell checks, each group’s numerical audit, and the integrated browser result. All **98,544 numeric native cache values with resolvable worksheet references** match their source cells. Known broken references and two legitimate literal series are accounted for separately.

## Structure

```text
index.html                 Complete offline deliverable
src/index.html             Page template
src/styles.css             Responsive interface
src/app.js                 Navigation, plotting, inspection and exports
src/charts/*.json          Independently generated figure specifications
scripts/chart_common.py    OOXML chart parser and shared conversion
scripts/build_group_*.py   Native chart adjustments and checks
scripts/build_supplemental*.py  Image-backed numerical reconstructions
scripts/rebuild.py         Parallel reproducible build
data/sources/              Original AEMO files and checksums
data/analysis/             Inventory, images and validation evidence
vendor/                   Pinned Plotly.js and license
```
