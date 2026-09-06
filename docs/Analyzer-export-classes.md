# Classes needed for the Analyzer export feature

The Analyzer's "Excel Workbook" and "JSON" export options (added to the stock
Export dropdown) are implemented across these classes:

- **`BIExport.UI.Analyzer`** - subclasses the stock `%DeepSee.UI.Analyzer` page.
  DOM-patches the Export dropdown to add the two new items, and builds the
  export URL each one navigates to (`buildExportURL()` in this class). Also
  calls `BIExport.XLSXWriter.IsAvailable()` (via its own `IsXLSXAvailable`
  ZenMethod) to gray out "Excel Workbook" when Python/openpyxl aren't set up.

- **`BIExport.UI.XLSXPage`** - the page "Excel Workbook" navigates to. A plain
  `%CSP.Page` (not `%CSP.REST`) living in the same CSP application as the
  Analyzer, so the browser's existing login session is reused. Calls
  `BIExport.XLSXWriter.IsAvailable()` again server-side (belt-and-suspenders -
  the Analyzer's check only controls whether the link is clickable) before
  building the workbook.

- **`BIExport.UI.JSONPage`** - the page "JSON" navigates to. Same
  same-CSP-application reasoning as `XLSXPage`, no Python dependency.

- **`BIExport.MDXExport`** - shared by both export pages. Runs the pivot's MDX
  and shapes the result into the row/column/data structure the writers below
  consume (`BuildExportData()`), and parses the `FILTERNAMES`/`FILTERVALUES`
  query params (`%ParseFilterPairs()`).

- **`BIExport.XLSXWriter`** - turns `MDXExport`'s data into the actual `.xlsx`
  bytes (`BuildXLSX()`), via IRIS Embedded Python + openpyxl. Also owns the
  `IsAvailable()` check both `Analyzer` and `XLSXPage` call.

JSON export has no equivalent writer class - `JSONPage` calls `%ToJSON()`
directly on `MDXExport`'s output.

## Request flow

```
Analyzer dropdown click
  -> exportXLSX() / exportJSON()          (BIExport.UI.Analyzer, client JS)
  -> GET BIExport.UI.XLSXPage.cls / BIExport.UI.JSONPage.cls
       -> BIExport.MDXExport.BuildExportData()
       -> BIExport.XLSXWriter.BuildXLSX()  (XLSX only)
```

Gray-out check (page load, before any export click):

```
Analyzer.onloadHandler()
  -> IsXLSXAvailable() ZenMethod          (BIExport.UI.Analyzer)
  -> BIExport.XLSXWriter.IsAvailable()
```
