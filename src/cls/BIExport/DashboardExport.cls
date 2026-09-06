/// Pure ObjectScript (no Python): resolves every pivot widget on a saved dashboard into
/// the same normalized export structure <class>BIExport.MDXExport</class> already
/// produces for a single pivot, so <class>BIExport.XLSXWriter</class>'s
/// <method>BuildMultiSheetXLSX</method> can turn a whole dashboard into one workbook -
/// one sheet per pivot widget - without any of the cellset-flattening logic being
/// duplicated. <method>BuildDashboardExportData</method> is this class's only public
/// entry point; everything else is a private helper for it.
///
/// Resolving a saved pivot into a live, queryable <class>%DeepSee.Component.pivotTable</class>
/// uses the exact same <method>%OpenFolderItem</method> + <method>%CopyToComponent</method>
/// sequence <class>%DeepSee.Component.Widget.pivot</class>'s own <method>%LoadPivotInfo</method>
/// uses (confirmed live by dumping that method's compiled source) - but here the
/// <class>%DeepSee.Component.pivotTable</class> is a plain, page-less <method>%New</method>()
/// object, never attached to any Zen component tree. That side-steps the live-component
/// property-read trap <class>BIExport.UI.Widget.pivot</class> had to work around (reading
/// a property off a component already wired into a rendering page's tree can throw
/// <code><OBJECT DISPATCH> ... must be MultiDimensional</code>) - it doesn't apply here
/// because this object is never attached to a page in the first place. Confirmed live
/// with a throwaway script against a real fixture pivot before writing this class:
/// %New() -> %CopyToComponent() -> GetCurrentQueryText("export") -> MDXExport.BuildExportData()
/// round-tripped cleanly end to end.
///
/// v1 limitation (also called out in README.md/CLAUDE.md): every widget is exported
/// using its SAVED pivot state - whatever <property>dataSource</property> points at,
/// as last saved - not whatever the CURRENT on-screen dashboard viewer session has live
/// (unsaved filter changes, unsaved widget option tweaks). There's no browser session
/// available to a plain server-side classmethod like this one, so live on-screen state
/// genuinely isn't recoverable here - contrast with <class>BIExport.UI.Widget.pivot</class>'s
/// per-widget export, which DOES reflect live state because it runs client-side against
/// the actual rendered <class>%DeepSee.Component.pivotTable</class>.
///
/// One more consequence of having no live pivot table widget to ask: row-header column
/// captions (e.g. "YearSold", "Product Category" - see MDXExport.cls's own doc comment
/// and CLAUDE.md gotcha #4) come from the pivot table WIDGET's own
/// <method>getRowCaptions</method>(), which itself falls back to whatever the last
/// server-rendered response happened to include - there is no equivalent available to a
/// plain ObjectScript classmethod with no rendered widget at all. This class simply
/// doesn't pass a <var>rowCaptions</var> option, so every widget's sheet gets the same
/// generic "Row 1"/"Row 2" fallback labels <class>MDXExport</class> already uses whenever
/// that option is omitted - not a new limitation, the existing documented fallback.
Class BIExport.DashboardExport Extends %RegisteredObject
{

/// Opens the dashboard named by <var>pDashboardFullName</var> (its full library path,
/// e.g. "BIExport Test Dashboard.dashboard") and returns a <class>%DynamicObject</class>
/// shaped like:
/// <example>
/// {
///   "title": "...",            // the dashboard's own title property
///   "dashboardName": "...",    // its fullName, echoed back
///   "sheets": [
///     {"widgetName":"Widget1", "title":"...", "data": { ...BuildExportData() shape... }},
///     {"widgetName":"Widget2", "title":"...", "error": "..."}   // widget that failed
///   ],
///   "skipped": [
///     {"widgetName":"Widget3", "title":"...", "type":"..."}     // non-pivot widget
///   ]
/// }
/// </example>
/// One sheet entry per pivot-family widget (see <method>%IsPivotWidgetType</method>) -
/// present even if that widget's own export failed, with an "error" field instead of
/// "data" so one bad widget can't take down the whole export. Non-pivot-family widgets
/// (chart-catalog/KPI/text) never get a sheet entry at all; they're listed separately
/// under "skipped" purely so <class>BIExport.XLSXWriter</class> can mention them by name
/// on the combined Info sheet, per the plan's "clear message, not a bug" call.
///
/// <var>pStatus</var> is $$$OK unless the dashboard itself couldn't be opened at all -
/// an individual widget failure is recorded on its own sheet entry instead, never surfaced
/// through <var>pStatus</var>.
ClassMethod BuildDashboardExportData(pDashboardFullName As %String, Output pStatus As %Status) As %DynamicObject
{
    Set pStatus = $$$OK
    Set tResult = {}
    Try {
        Set tDash = ##class(%DeepSee.Dashboard.Utils).%OpenDashboard(pDashboardFullName, .pStatus)
        If $$$ISERR(pStatus) Quit
        If '$IsObject(tDash) {
            Set pStatus = $$$ERROR($$$GeneralError, "Dashboard not found: "_pDashboardFullName)
            Quit
        }

        Set tResult.title = tDash.title
        Set tResult.dashboardName = tDash.fullName

        Set tSheets = []
        Set tSkipped = []
        For i=1:1:tDash.widgets.Count() {
            Set tWidget = tDash.widgets.GetAt(i)
            If '..%IsPivotWidgetType(tWidget.type) {
                Do tSkipped.%Push({"widgetName": (tWidget.name), "title": (tWidget.title), "type": (tWidget.type)})
                Continue
            }
            Do tSheets.%Push(..%BuildWidgetSheet(tWidget))
        }
        Set tResult.sheets = tSheets
        Set tResult.skipped = tSkipped
    }
    Catch ex {
        Set pStatus = ex.AsStatus()
    }
    Quit tResult
}

/// True if <var>pType</var> (a widget's own <property>type</property>) is something
/// resolvable the same way <class>%DeepSee.Dashboard.Widget::%CreateVisualWidget</class>
/// resolves it into a real class (blank/"pivot" prefixed to
/// "%DeepSee.Component.Widget.pivot"; anything containing a dot used as a literal class
/// name - same prefixing rule <class>BIExport.UI.Widget.pivot</class>'s own doc comment
/// already documents) AND that resolved class actually extends
/// <class>%DeepSee.Component.Widget.pivot</class>. More general than
/// <class>BIExport.DashboardTools</class>'s own <method>%IsPivotFamily</method> (which
/// only needs to recognize its own conversion source/target literal strings) - this one
/// needs to correctly classify ANY pivot-widget subclass a dashboard might already be
/// using, not just the one this package ships.
ClassMethod %IsPivotWidgetType(pType As %String) As %Boolean [ Private ]
{
    If (pType = "") || (pType = "pivot") Quit 1
    Set tClass = $Select(pType[".":pType, 1:"%DeepSee.Component.Widget."_pType)
    Quit ##class(%Dictionary.CompiledClass).%ExistsId(tClass) && $ClassMethod(tClass, "%Extends", "%DeepSee.Component.Widget.pivot")
}

/// Resolves one pivot-family widget's saved pivot into a sheet entry - see
/// <method>BuildDashboardExportData</method>'s doc comment for the shape. Never throws;
/// any failure (missing/bad saved pivot, MDX error) is caught and recorded as this
/// sheet's own "error" field instead.
ClassMethod %BuildWidgetSheet(pWidget As %DeepSee.Dashboard.Widget) As %DynamicObject [ Private ]
{
    Set tSheet = {}
    Set tSheet.widgetName = pWidget.name
    Set tSheet.title = $Select(pWidget.title'="":pWidget.title, 1:pWidget.name)
    Try {
        Set tPivotName = pWidget.dataSource
        Set tExt = $Piece(tPivotName,".",$Length(tPivotName,"."))
        If (tPivotName = "") || (tExt '= "pivot") {
            // KPI ("*.kpi"), metric, worksheet, or no data source at all (e.g. a chart
            // pulling from another widget) - %IsPivotWidgetType() only looks at the
            // widget's rendering CLASS, not what it's actually pointed at, so this can
            // still happen for a widget using BIExport.UI.Widget.pivot (or stock "pivot")
            // whose dataSource isn't a plain saved pivot. Same non-fatal "error" field
            // treatment as any other per-widget failure - just a more specific message.
            Set tSheet.error = "Widget's data source is not a saved pivot (dataSource="""_tPivotName_"""); skipping its data."
            Quit
        }
        If '##class(%DeepSee.UserLibrary.Utils).%FolderItemExists(tPivotName) {
            Set tSheet.error = "Saved pivot not found: "_tPivotName
            Quit
        }

        Set tPivot = ##class(%DeepSee.UserLibrary.Utils).%OpenFolderItem(tPivotName,.tSC)
        If $$$ISERR(tSC) Throw ##class(%Exception.StatusException).CreateFromStatus(tSC)

        Set tTable = ##class(%DeepSee.Component.pivotTable).%New()
        Set tSC = tPivot.%CopyToComponent(tTable)
        If $$$ISERR(tSC) Throw ##class(%Exception.StatusException).CreateFromStatus(tSC)

        Set tMDX = tTable.GetCurrentQueryText("export")
        // GetCurrentQueryText() never throws - on failure it returns the error text
        // itself as the "MDX" (see its own compiled source), so a failure has to be
        // caught by inspecting the string rather than a status/exception.
        If ($Extract(tMDX,1,6) = "ERROR:") {
            Set tSheet.error = tMDX
            Quit
        }

        Set tOptions = {}
        Set tOptions.title = tSheet.title
        Set tOptions.subtitle = tTable.printSubtitle
        Set tOptions.cube = tTable.cubeName
        Set tOptions.listing = tTable.listing
        // Crossed rowTotals/columnTotals(+Agg) mapping - see CLAUDE.md gotcha #8 and
        // BIExport.UI.Widget.pivot's buildExportURL() doc comment: table.rowTotals
        // renders a Totals ROW in the widget (-> our COLUMNTOTALS), table.columnTotals
        // renders a Totals COLUMN (-> our ROWTOTALS). Same crossing, just read straight
        // off the saved pivot's own properties instead of a live client-side table.
        Set tOptions.rowTotals = tTable.columnTotals
        Set tOptions.columnTotals = tTable.rowTotals
        Set tOptions.rowTotalAgg = $Select(tTable.columnTotalAgg'="":tTable.columnTotalAgg, 1:"sum")
        Set tOptions.columnTotalAgg = $Select(tTable.rowTotalAgg'="":tTable.rowTotalAgg, 1:"sum")

        Set tSheet.data = ##class(BIExport.MDXExport).BuildExportData(tMDX, tOptions)
    }
    Catch ex {
        Set tSheet.error = ex.DisplayString()
    }
    Quit tSheet
}

}
