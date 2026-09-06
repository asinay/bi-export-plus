/// Subclass of the stock dashboard viewer page that adds a real, whole-dashboard
/// multi-sheet ".xlsx" export - one workbook with one sheet per pivot widget (see
/// <class>BIExport.DashboardExport</class>/<class>BIExport.XLSXWriter.BuildMultiSheetXLSX</class>) -
/// next to the per-widget "Excel Workbook"/"JSON" options <class>BIExport.UI.Widget.pivot</class>
/// already adds to each widget's own export dropdown.
///
/// Reachable at the same CSP application as the stock page, just under this class's
/// name instead of "_DeepSee.UserPortal.DashboardViewer.zen", e.g.
/// /csp/<namespace>/BIExport.UI.DashboardViewer.cls?DASHBOARD=... - point any existing
/// shortcuts/links at this URL instead of the stock one to pick up the new option.
///
/// DEVIATION FROM PLAN: the plan's sketch assumed an existing page-level "icon stack"
/// to inject a new standalone toolbar icon next to - confirmed LIVE (dumping this page's
/// compiled XData layout, then loading the real page with Playwright and screenshotting
/// it) that no such thing exists. The only page-level chrome here is the "Menu" dropdown
/// button (an absolutely-positioned <class>%ZEN.Component.menu</class> the plan already
/// separately ruled out DOM-patching into - confirmed right, its own contextMenu content
/// is server-sized) and a plain-text small-menu row - "Home | About | Save | Logout" -
/// built by the overridable <method>%OnGetSmallMenu</method>(*pMenu) callback (every
/// widget's OWN icon row - table/chart/dim-tree/print/export/search icons - is a
/// per-widget thing already handled by <class>BIExport.UI.Widget.pivot</class>, not a
/// page-level toolbar). Since there's no icon markup to model a new icon on, and
/// <method>%OnGetSmallMenu</method> is a small, officially overridable server-side
/// callback (confirmed live by dumping its compiled source: it just returns
/// $ListBuild(caption,title,script) rows for the caller to render as text links) rather
/// than a DOM structure to patch, this class adds "Export Dashboard" there instead -
/// safer than a client-side DOM patch (no risk of the live-component property-read trap
/// <class>BIExport.UI.Widget.pivot</class> had to work around, since nothing here reads
/// a property off an already-rendered component at all) and consistent with how
/// <property>ANALYZERURL</property>/<property>viewerURL</property> are also plain
/// server-side overrides rather than client patches.
///
/// GOTCHA hit writing this class (new, not previously documented): the stock
/// <method>%OnGetSmallMenu</method>'s formal spec, as dumped from the compiled class,
/// shows its one parameter's passing mechanism as "*" (Output) - but writing the
/// override's own formal spec using that same bare <code>*pMenu</code> shorthand made
/// THIS ENTIRE CLASS fail to compile with a generic "ERROR #5559 ... non-matching {}
/// or () ... or a missing ; character" parse error (confirmed by bisection: an
/// otherwise-empty method with nothing but <code>Quit</code> in its body still failed
/// with the exact same error, and swapping only the formal spec to the equivalent
/// <code>Output pMenu</code> keyword form fixed it immediately, with no other change).
/// Lesson: <code>*name</code> is how a compiled FormalSpec DISPLAYS an Output parameter
/// after the fact, not necessarily valid UDL SOURCE syntax for declaring one in every
/// context - prefer the explicit <code>Output</code>/<code>ByRef</code> keywords when
/// writing a method by hand instead of guessing from a dumped signature.
Class BIExport.UI.DashboardViewer Extends %DeepSee.UserPortal.DashboardViewer
{

/// The stock page's own ancestor (<class>%DeepSee.UserPortal.standardPage</class>)
/// seeds the client-side <property>viewerURL</property> from
/// <method>%DeepSee.UserPortal.Utils::%GetViewerURL</method>() - a hardcoded literal,
/// "_DeepSee.UserPortal.DashboardViewer.zen" - inside ITS OWN <method>%OnAfterCreatePage</method>
/// (confirmed live by dumping both methods' compiled source). Every internal nav action
/// (Open/New/Save/Save-As, drill-down, etc.) builds its redirect URL from that property,
/// so left unoverridden, all of them would bounce back to the stock page - same
/// "%OnAfterCreatePage sets a URL a subclass has to reassign" shape already documented
/// for <property>ANALYZERURL</property> on <class>BIExport.UI.Analyzer</class>, just via
/// a plain property post-assignment here instead of a class <parameter>Parameter</parameter>,
/// since <property>viewerURL</property> isn't parameter-driven.
Method %OnAfterCreatePage() As %Status
{
    Set tSC = ##super()
    If $$$ISERR(tSC) Quit tSC
    Set ..viewerURL = ..%ClassName(1) _ ".cls"
    Quit tSC
}

/// Stock <class>%DeepSee.UserPortal.DashboardViewer</class> already overrides this
/// (fully replacing, not extending, its own ancestor's Home/About/Logout list - see its
/// compiled source), so calling <method>##super</method> here inherits ITS list
/// (Home/About/[Save if writable]/Logout) - this just appends one more entry after it.
/// See this class's own doc comment for why <var>pMenu</var> is declared
/// <code>Output pMenu</code> rather than the <code>*pMenu</code> shorthand the stock
/// method's own dumped signature shows.
///
/// Plain string literals rather than <code>$$$Text()</code> (which every stock entry
/// in this same list uses) - confirmed live that macro tries to auto-register its
/// string into the "%DeepSee" message domain at COMPILE time (this class's inherited
/// <parameter>DOMAIN</parameter>), which is product-owned and not writable from here:
/// compiling failed with <code>&lt;PROTECT&gt;</code> deep inside
/// <code>macroText^%occMessages</code>. Not worth adding a whole new message domain
/// just for one menu item's localization.
Method %OnGetSmallMenu(Output pMenu)
{
    Do ##super(.pMenu)
    Set tNext = $Order(pMenu(""),-1) + 1
    // Whole-dashboard export needs Python too (BuildMultiSheetXLSX, same as the per-widget
    // button) - gated here the same way as BIExport.UI.Widget.pivot's dropdown option, since
    // this menu is also built entirely server-side. Confirmed live (dumping %DrawSmallMenu's
    // compiled source) that this row shape - $ListBuild(caption,title,script) - has no visual
    // "disabled" concept: every row renders with the identical "portalTitleLink" CSS class
    // regardless of what's in it, and an empty script just makes the link's onclick a no-op
    // ("return false;", no JS error) - so this can't be grayed out the way the pivot widget's
    // dropdown option field (raw HTML) can. Still shown (not omitted) with a distinct caption
    // and the reason as its hover tooltip, since a click-time landing on a friendly error is
    // strictly worse discoverability than a link that's visibly there but explains itself on
    // hover and does nothing when clicked.
    //
    // %DrawSmallMenu HTML-escapes the caption (position 1) itself but writes the title/tooltip
    // (position 2) verbatim - confirmed the same way - so tXlsxReason is escaped here before
    // going into that slot.
    If ##class(BIExport.XLSXWriter).IsAvailable(.tXlsxReason) {
        Set pMenu(tNext) = $ListBuild("Export Dashboard", "Export every pivot widget on this dashboard into one multi-sheet Excel workbook", "zenPage.exportDashboardXLSX();")
    } Else {
        Set pMenu(tNext) = $ListBuild("Export Dashboard (unavailable)", $ZCVT(tXlsxReason,"O","HTML"), "")
    }
}

/// Opens the whole-dashboard multi-sheet .xlsx export for the CURRENT dashboard (whatever
/// <property>dashboardName</property> - the same client-mirrored property every other
/// nav action on this page already reads - happens to be right now). Same
/// same-CSP-application/window.open() rationale as
/// <class>BIExport.UI.Widget.pivot</class>/<class>BIExport.UI.Analyzer</class>'s own
/// export buttons: BIExport.UI.DashboardXLSXPage is a plain %CSP.Page in this SAME
/// application, so the browser's already-open session covers it with no extra login.
ClientMethod exportDashboardXLSX() [ Language = javascript ]
{
    if (!this.dashboardName) {
        alert('No dashboard is currently loaded.');
        return;
    }
    var url = 'BIExport.UI.DashboardXLSXPage.cls?DASHBOARD=' + encodeURIComponent(this.dashboardName);
    window.open(zenLink(url), 'dashboardxlsx', '');
}

}
