# bi-export-plus

Real, binary `.xlsx` (and typed JSON) exports for the InterSystems IRIS BI/DeepSee
Analyzer — no more opening "Export to Excel" and getting an HTML table wearing an
`.xls` extension.

- **Excel Workbook** and **JSON** buttons appear right in the Analyzer's own Export
  dropdown, next to the stock Excel/CSV options.
- A standalone REST API (`/csp/bi-export/xlsx`, `/csp/bi-export/json`) too, for
  pulling the same exports from scripts/automation without opening the Analyzer.

**New here?** → **[Try it in 2 minutes](QUICKSTART.md)**. Everything below is the
full reference: every install option, every request field, every limitation.

## Requirements

- InterSystems IRIS with **Embedded Python** enabled.
- The Python package [`openpyxl`](https://pypi.org/project/openpyxl/) importable
  from IRIS's embedded Python (needed for the `.xlsx` path only; JSON export
  needs neither Python nor openpyxl).

Install openpyxl into IRIS's embedded Python environment, e.g.:

```bash
<iris-install-dir>/bin/irispython -m pip install openpyxl
```

Verify it from an IRIS terminal in any namespace:

```objectscript
Do ##class(%SYS.Python).Import("openpyxl")
```

If that errors, fix the Python/openpyxl setup before installing this module -
the xlsx export will fail at runtime otherwise.

## Installation

**Recommended: InterSystems Package Manager (IPM/ZPM).** If this module has been
published to the IPM registry:

```objectscript
zpm "install bi-export-plus"
```

To install straight from a local clone of this repository instead (e.g. before
it's published, or to run a modified copy):

```bash
git clone <this-repo-url>
```

```objectscript
zpm "load /path/to/bi-export-plus"
```

Point this at the **directory that contains `module.xml`**, not at the
`module.xml` file itself - e.g. on Windows:

```
zpm:USER>load C:\path\to\bi-export-plus
```

IPM reads [module.xml](module.xml), loads the classes under [src/](src), and
registers the `/csp/bi-export` web application for you. Take note of the
`AfterInstallMessage` it prints at the end - it points you at the two usage
options below.

<details>
<summary><strong>Other ways to install</strong> (no IPM, or just the packaged classes XML)</summary>

### Option B: Manual install (no IPM)

1. Load the classes into your target namespace from an IRIS terminal:

   ```objectscript
   Set path = "/path/to/bi-export-plus/src/cls"
   Do $system.OBJ.LoadDir(path, "ck", , 1)
   ```

   (`"ck"` = compile, keep the source-controlled state consistent; the trailing
   `1` recompiles even if the class already exists.)

2. Create the REST web application manually (IPM normally does this from the
   `<WebApplication>` block in [module.xml](module.xml)) - in the Management
   Portal: **System Administration > Security > Applications > Web
   Applications > New Web Application**, with:

   | Setting | Value |
   |---|---|
   | Name | `/csp/bi-export` |
   | Namespace | your target namespace |
   | Dispatch class | `BIExport.REST` |
   | Enable Password authentication | Yes |
   | Enable Unauthenticated access | No |

   This step is only needed for the standalone REST API (Option B under
   "Usage" below); the in-Analyzer export buttons work without it.

3. Confirm the openpyxl import (see Requirements above) - manual installs skip
   IPM's install-time checks, so this is easy to miss.

### Option C: Import the packaged classes XML

[dist/BIExport-classes.xml](dist/BIExport-classes.xml) is a single IRIS class
export containing every class in this module - the quickest way to get it
into an existing instance without cloning the repo or using IPM.

Either drag-and-drop it onto **Management Portal > System Explorer >
Classes > Import**, or from an IRIS terminal in your target namespace:

```objectscript
Do $system.OBJ.Load("/path/to/bi-export-plus/dist/BIExport-classes.xml", "ck")
```

This does **not** create the `/csp/bi-export` web application - it's just the
classes. That's enough for the in-Analyzer export buttons (Option A under
"Usage" below), but if you also want the standalone REST API, create the web
application manually per step 2 of Option B above.

</details>

## Usage

### Option A: In-Analyzer export buttons

Open your pivots via `BIExport.UI.Analyzer.cls` instead of the stock
`_DeepSee.UI.Analyzer.zen` - same query params, just a different class name:

```
/csp/<namespace>/BIExport.UI.Analyzer.cls?CUBE=...
```

The page's Export dropdown gains **Excel Workbook** and **JSON** entries next
to the stock Excel/CSV ones.

<details>
<summary>Add a Favorites bookmark instead of changing a link's URL</summary>

Rather than changing an existing link/shortcut's URL, add a normal Portal
**Favorite** that opens a saved pivot through `BIExport.UI.Analyzer.cls`
directly - it shows up in the User Portal's Favorites sidebar/cover shelf
like any other bookmark:

```objectscript
Do ##class(BIExport.Favorites).AddAnalyzerFavorite("<folder>/<pivot name>.pivot")
```

The one required argument is the pivot's full library path (`fullName` in
the Portal's pivot list). Optional second/third arguments choose the
bookmark's folder (default `"BIExport"`) and display title (defaults to the
pivot's own name). Run once per pivot, per user - Favorites are personal to
each user, there's no bulk/all-users mode. `RemoveAnalyzerFavorite(pFolderName,
pTitle)` undoes it.

</details>

### Option B: Standalone REST API

`POST` an MDX query to `/csp/bi-export/xlsx` or `/csp/bi-export/json`:

```bash
curl -u _system:PASSWORD -X POST \
  http://localhost:52773/csp/bi-export/xlsx \
  -H "Content-Type: application/json" \
  -d '{
    "mdx": "SELECT [Measures].[Amount] ON 0 FROM [HoleFoods Sales]",
    "title": "Sales Report",
    "cube": "HoleFoods Sales"
  }' \
  -o report.xlsx
```

Supported request body fields:

| Field | Required | Description |
|---|---|---|
| `mdx` | yes | MDX query text |
| `filename` | no | Defaults to `<title>.xlsx`/`.json`, or `<cube>_<timestamp>` if no title |
| `title`, `subtitle` | no | Shown on the Info sheet |
| `cube` | no | Shown on the Info sheet, used in the default filename |
| `listing` | no | Listing name, for drillthrough/listing MDX |
| `rowTotals` | no | `true` to include a "Total" column (one value per row, summed across columns). Independent of `columnTotals` — a pivot can have either, both, or neither. |
| `columnTotals` | no | `true` to include a "Total" row (one value per column, summed down rows). Independent of `rowTotals`. |
| `rowTotalAgg` | no | Aggregation for `rowTotals` — `"sum"` (default), `"count"`, `"min"`, `"max"`, `"avg"`, `"pct"` (% of total) |
| `columnTotalAgg` | no | Aggregation for `columnTotals`, same values as `rowTotalAgg` |
| `filters` | no | `[{"name": "...", "value": "..."}]` recap block on the Info sheet |
| `rowCaptions` | no | Row axis dimension captions (outermost first), for the row-header columns |

See sample output in [samples/](samples).

## Known limitations

- **Totals across mixed measures.** If a row/column axis mixes different
  measures (e.g. "Amount Sold" and "Units Sold" side by side), `rowTotals`/
  `columnTotals` aggregate across them regardless of measure, mixing dollars
  and units - this matches the stock Analyzer's own CSV/Excel exports, which
  call the same underlying IRIS API. Totals are only meaningful when an axis
  is entirely one measure.
- **Nested column headers aren't Excel Tables.** A pivot with a dimension on
  Columns (e.g. Channel x measure) needs a true 2+ row merged header, which
  Excel's Table feature can't have (exactly one header row allowed) - that
  export keeps the merged header and skips the Table (no autofilter/banded
  rows), with a note on the Info sheet explaining why.

## License

MIT - see [LICENSE](LICENSE).
