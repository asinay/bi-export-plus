# Classes needed for the Dashboard export features

Two related features, both covered in README's "Dashboard export" section:
per-widget export buttons on an individual pivot widget, and a whole-dashboard
multi-sheet export. They share some classes but have separate entry points.

## Feature 1: Per-widget export buttons

Same "Excel Workbook"/"JSON" options as the Analyzer, but on one dashboard
widget's own Export dropdown.

- **`BIExport.UI.Widget.pivot`** - subclass of the stock
  `%DeepSee.Component.Widget.pivot`. A widget opts in by setting its
  `type` to this class (see `DashboardTools` below). Adds the dropdown options
  two different ways depending on the widget's `csv` property (`%OnCreateControls`
  server-side rebuild for the single-icon case, `addExtraExportOptions()`
  client-side DOM patch for the existing-dropdown case) - both gated on
  `BIExport.XLSXWriter.IsAvailable()` the same way `BIExport.UI.Analyzer` is.
  Its `buildExportURL()` is a widget-context port of the Analyzer's own method.

- **`BIExport.DashboardTools`** - rollout helper, not part of the request path
  at runtime. `ConvertWidget()`/`ConvertAllWidgets()` flip an existing widget's
  `type` to `BIExport.UI.Widget.pivot` (and `RevertWidget()` undoes it) so
  existing dashboards don't need hand-edited XML. New dashboards can set the
  widget type directly in the Dashboard Editor instead and skip this class
  entirely.

- **`BIExport.UI.XLSXPage`**, **`BIExport.UI.JSONPage`**, **`BIExport.MDXExport`**,
  **`BIExport.XLSXWriter`** - the same single-pivot export pipeline the
  Analyzer uses (see [Analyzer-export-classes.md](Analyzer-export-classes.md)).
  A widget's export button ends up on the exact same pages, just with the MDX
  sourced from the widget's live `pivotTable` component instead of the
  Analyzer page's own.

### Request flow

```
Widget dropdown click
  -> exportXLSX() / exportJSON()          (BIExport.UI.Widget.pivot, client JS)
  -> GET BIExport.UI.XLSXPage.cls / BIExport.UI.JSONPage.cls
       -> BIExport.MDXExport.BuildExportData()
       -> BIExport.XLSXWriter.BuildXLSX()  (XLSX only)
```

## Feature 2: Whole-dashboard export

One `.xlsx` workbook with one sheet per pivot widget on a saved dashboard,
via a new **Export Dashboard** entry in the dashboard viewer's small-menu row.

- **`BIExport.UI.DashboardViewer`** - subclass of the stock
  `%DeepSee.UserPortal.DashboardViewer`. Overrides `%OnGetSmallMenu()` to add
  "Export Dashboard" (gated on `BIExport.XLSXWriter.IsAvailable()` - shown
  either way, since this menu shape has no visual "disabled" state, just a
  no-op click and an explanatory tooltip when unavailable). Its
  `exportDashboardXLSX()` opens `DashboardXLSXPage` for the current
  `dashboardName`.

- **`BIExport.UI.DashboardXLSXPage`** - the page "Export Dashboard" opens. A
  plain `%CSP.Page` in the same CSP application as `DashboardViewer`, mirroring
  `BIExport.UI.XLSXPage`'s reasoning exactly (shared login session, no separate
  REST application). Reads the `DASHBOARD` query param, and 503s the same way
  `XLSXPage` does when Python/openpyxl aren't available.

- **`BIExport.DashboardExport`** - pure ObjectScript, resolves every pivot-family
  widget on the named dashboard into `MDXExport`'s normalized structure
  (`BuildDashboardExportData()`), by opening each widget's *saved* pivot into a
  page-less `%DeepSee.Component.pivotTable` (never attached to a real page - side-
  steps the live-component property-read trap `BIExport.UI.Widget.pivot` has to
  work around). A widget that isn't pivot-family is listed under `"skipped"`;
  a pivot widget whose own export fails gets an `"error"` field on its sheet
  entry instead of failing the whole export.

- **`BIExport.MDXExport`** - same per-widget data shaping `DashboardExport`
  calls into, one widget at a time.

- **`BIExport.XLSXWriter`** - `BuildMultiSheetXLSX()` turns `DashboardExport`'s
  multi-sheet structure into the actual workbook (one sheet per widget, plus
  an Info sheet listing any skipped/errored widgets). Also the source of the
  `IsAvailable()` check both `DashboardViewer` and `DashboardXLSXPage` use.

### Request flow

```
"Export Dashboard" click
  -> exportDashboardXLSX()                 (BIExport.UI.DashboardViewer, client JS)
  -> GET BIExport.UI.DashboardXLSXPage.cls?DASHBOARD=...
       -> BIExport.DashboardExport.BuildDashboardExportData()
            -> BIExport.MDXExport.BuildExportData()   (per widget)
       -> BIExport.XLSXWriter.BuildMultiSheetXLSX()
```

### Known limitation shared by this feature (not Feature 1)

Whole-dashboard export reads each widget's **saved** pivot state, not whatever
the dashboard viewer session currently has on-screen (unsaved filter/option
changes) - there's no browser session available to a server-side classmethod
like `DashboardExport`. Row headers also fall back to generic "Row 1"/"Row 2"
labels, since there's no live pivot table widget to ask for row captions. The
per-widget buttons in Feature 1 don't have either limitation - they run
against the widget's live, on-screen state. See README's "Known limitations"
section.
