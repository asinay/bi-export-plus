# bi-export-plus

Real, binary `.xlsx` (and typed JSON) exports for the InterSystems IRIS BI/DeepSee
Analyzer, replacing the legacy HTML-saved-as-`.xls` "Export to Excel" behavior.

- Adds **Excel Workbook** and **JSON** items to the Analyzer's own Export dropdown
  (alongside the stock Excel/CSV options).
- Adds a standalone REST API (`/csp/bi-export/xlsx`, `/csp/bi-export/json`) so
  external tools/scripts can pull the same exports without opening the Analyzer.

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

### Option A: InterSystems Package Manager (IPM/ZPM)

If this module has been published to the IPM registry:

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
zpm:USER>load C:\AI_Apps\bi-export-plus
```

IPM reads [module.xml](module.xml), loads the classes under [src/](src), and
registers the `/csp/bi-export` web application for you. Take note of the
`AfterInstallMessage` it prints at the end - it points you at the two usage
options below.

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
Do $system.OBJ.Load("C:\AI_Apps\bi-export-plus\dist\BIExport-classes.xml", "ck")
```

This does **not** create the `/csp/bi-export` web application - it's just the
classes. That's enough for the in-Analyzer export buttons (Option A under
"Usage" below), but if you also want the standalone REST API, create the web
application manually per step 2 of Option B above.

## Usage

### Option A: In-Analyzer export buttons

Open your pivots via `BIExport.UI.Analyzer.cls` instead of the stock
`_DeepSee.UI.Analyzer.zen` - same query params, just a different class name:

```
/csp/<namespace>/BIExport.UI.Analyzer.cls?CUBE=...
```

The page's Export dropdown gains **Excel Workbook** and **JSON** entries next
to the stock Excel/CSV ones.

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
| `totals` | no | `true` to include row/column totals |
| `filters` | no | `[{"name": "...", "value": "..."}]` recap block on the Info sheet |
| `rowCaptions` | no | Row axis dimension captions (outermost first), for the row-header columns |

See sample output in [samples/](samples).

## License

MIT - see [LICENSE](LICENSE).
