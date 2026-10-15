/// Subclass of the stock dashboard pivot-table widget that adds real "Excel Workbook"
/// (.xlsx) and "JSON" export options next to the existing "Excel"/"CSV"/"CSV with
/// Headers" ones - the same idea as <class>BIExport.UI.Analyzer</class>, but for
/// dashboard widgets rather than the full-page Analyzer.
///
/// A dashboard widget's rendering class is swappable via its own definition: set the
/// widget's <property>type</property> to "BIExport.UI.Widget.pivot" (e.g. via
/// <class>BIExport.DashboardTools</class>) and <method>%DeepSee.Dashboard.Widget::%CreateVisualWidget</method>
/// instantiates it directly in place of the stock "pivot" type - confirmed live by
/// dumping that method's compiled source, same technique already used to confirm
/// <property>ANALYZERURL</property> for <class>BIExport.UI.Analyzer</class>.
///
/// Two widget UI states exist, both handled here (confirmed live by dumping
/// <class>%DeepSee.Component.Widget.pivot</class>'s compiled <method>%OnCreateControls</method>):
///   - "csv" widget property off: the stock method renders a single bare
///     <img onclick="...exportExcel()"> icon - can't be DOM-patched (there's no
///     dropdown markup to patch at all), so this class's <method>%OnCreateControls</method>
///     override calls <method>%CreateExportDropdown</method> (a plain instance <class>Method</class>
///     on the stock class, callable from here) itself, after <method>##super</method> has
///     already rendered the bare icon, and swaps it out for a 3-option dropdown
///     (Excel / Excel Workbook / JSON) built from the exact same markup the "csv" widgets
///     already use natively.
///   - "csv" widget property on: the stock method already builds a
///     <caption>dsptIconDropdownMenu</caption> dropdown (Excel / CSV / CSV with Headers)
///     via that same <method>%CreateExportDropdown</method> call - left untouched
///     server-side and DOM-patched client-side instead, exactly like
///     <class>BIExport.UI.Analyzer</class> already does for the full-page Analyzer.
///
/// <method>addExtraExportOptions</method> (the client-side DOM patch) has to tell these
/// two cases apart so it never double-inserts into the dropdown this class's own
/// <method>%OnCreateControls</method> override already built server-side with the
/// extra options baked in - see that method's comment for how.
Class BIExport.UI.Widget.pivot Extends %DeepSee.Component.Widget.pivot
{

/// Calls <method>##super</method> first, unchanged, so every unrelated piece of stock
/// behavior (chart toggle, dimension list, PDF export, drill-down indicator, and the
/// native dropdown for "csv"=1 widgets) renders exactly as before. Only then, if the
/// stock method rendered the bare single-icon export button (export on at all - "excel"
/// or "csv" widget property truthy - but "csv" itself falsy), replaces that <img> child
/// in <var>pGroup</var> in place with an <class>%ZEN.Component.html</class> holding a
/// real dropdown - Excel (unchanged exportExcel()), "Excel Workbook" (exportXLSX()),
/// JSON (exportJSON()) - built via the stock class's own <method>%CreateExportDropdown</method>(id,
/// enabled, imgId, help, options) (confirmed live by dumping its compiled source: it
/// just assembles the same <caption>dsptIconDropdownMenu</caption> markup
/// <method>%OnCreateControls</method> already uses for "csv"=1 widgets - reusing it here
/// means both UI states end up with pixel-identical dropdown markup).
///
/// DEVIATION FROM PLAN: the plan sketch was to locate the stock single-icon button
/// afterwards by inspecting its rendered properties (e.g. its .onclick text, the same
/// marker the client-side DOM patch below uses on genuine "csv"=1 dropdowns). That does
/// not work here - confirmed live, with temporary debug logging, that reading ANY
/// property (.onclick, then .src as a second attempt) off a child already added to a
/// <var>pGroup</var> that is itself attached to a real, currently-rendering page throws
/// a live <code><OBJECT DISPATCH> Property '...' in class '%ZEN.Component.image' must
/// be MultiDimensional</code> error (silently swallowed by this method's own
/// Try/Catch below, which is why the symptom was "nothing happens" rather than a visible
/// error) - it did NOT reproduce in an ad-hoc Terminal test building a standalone,
/// page-less group/image pair, so it's specific to components already wired into a live
/// page's component tree, not a general property-access rule. Reading properties off a
/// *freshly built, not-yet-attached* object (see the options built below, and every
/// other class in this package) is fine - it's specifically already-attached children
/// that trip this.
///
/// Avoided entirely by never reading a property off the live child: the target button's
/// position within <var>pGroup.children</var> is instead computed arithmetically from
/// the exact same flags (chart toggle / dimension list / PDF) the stock method itself
/// uses to decide whether to add each preceding sibling - two of those three
/// (<property>chartToggle</property>, <property>hasDimTree</property>) are simply read
/// back off <var>..chartToggle</var>/<var>..hasDimTree</var> after <method>##super</method>
/// returns, since stock's own <method>%OnCreateControls</method> already set them as a
/// side effect on this same (our) object; the third (PDF button, "print" property) comes
/// from the same <var>tPValue</var> array built below. <method>%AddChildBefore</method>/
/// <method>%RemoveChild</method> only need the child's OREF (identity), never one of its
/// properties, so splicing it out this way is safe.
///
/// Left as a non-fatal <class>Try</class>/<class>Catch</class>: if this positional
/// arithmetic ever drifts out of sync with a future stock version's own ordering (stock
/// markup/logic changed shape), the widget still renders with its original single-icon
/// Excel-only export button rather than breaking the whole widget.
Method %OnCreateControls(pGroup As %ZEN.Component.group) As %Status
{
    Set tPrevCount = pGroup.children.Count()
    Set tSC = ##super(pGroup)
    If $$$ISERR(tSC) Quit tSC

    Try {
        Do ..%GetWidgetPropertyInfo(..subtype, .tList)
        Set n = $Order(tList(""))
        While (n '= "") {
            Set tName = $ListGet(tList(n), 1)
            Set tPValue(tName) = ..%definition.properties.GetAt(tName)
            Set:tPValue(tName)="" tPValue(tName) = $ListGet(tList(n), 2)
            Set n = $Order(tList(n))
        }

        Set tExcelOn = +$Get(tPValue("excel"))
        Set tCsvOn = +$Get(tPValue("csv"))

        // Only the single-icon case needs converting - genuine csv=1 widgets already
        // got a real dropdown from the stock method above, and the client-side
        // onloadHandler()/addExtraExportOptions() DOM-patches those instead.
        If (tExcelOn || tCsvOn) && ('tCsvOn) {
            // Positional lookup - see method doc comment for why. Mirrors the exact
            // order stock %OnCreateControls adds siblings in: chart toggle, then
            // dimension list, then PDF, then (here) export.
            Set tExportPos = tPrevCount
            Set:(..chartToggle="table")||(..chartToggle="chart") tExportPos = tExportPos + 1
            Set:..hasDimTree tExportPos = tExportPos + 1
            Set:(+$Get(tPValue("print"))) tExportPos = tExportPos + 1
            Set tExportPos = tExportPos + 1

            Set tImgChild = ""
            If (tExportPos > tPrevCount) && (tExportPos <= pGroup.children.Count()) {
                Set tImgChild = pGroup.children.GetAt(tExportPos)
            }

            If $IsObject(tImgChild) {
                Set tOptions = []

                Set tExcelOpt = {"name":"Excel"}
                Set tExcelOpt.click = "zenPage.getComponent("_..index_").exportExcel();"
                Set tExcelOpt.hover = $$$Text("Microsoft Excel (.xls)","%DeepSee")
                Do tOptions.%Push(tExcelOpt)

                // Grayed out (via a styled span in "name" - %CreateExportDropdown writes this
                // field verbatim as the <a>'s innerHTML, confirmed live by dumping its compiled
                // source, so this is the only per-option visual hook it actually offers - there
                // is no separate "disabled" flag on an individual option) and given no click
                // handler at all when BIExport.XLSXWriter.IsAvailable() reports IRIS Embedded
                // Python/openpyxl aren't available, with the reason as the hover tooltip -
                // guaranteed, not best-effort, since this whole option list is already built
                // server-side. %CreateExportDropdown writes "hover" verbatim into a title=""
                // attribute too (same dump), so it's HTML-escaped here - IsAvailable()'s own
                // reason text quotes CPF parameter values in double quotes, which would
                // otherwise prematurely close that attribute.
                Set tXlsxOpt = {"name":"Excel Workbook"}
                If ##class(BIExport.XLSXWriter).IsAvailable(.tXlsxReason) {
                    Set tXlsxOpt.click = "zenPage.getComponent("_..index_").exportXLSX();"
                    Set tXlsxOpt.hover = "Microsoft Excel Workbook (.xlsx)"
                } Else {
                    Set tXlsxOpt.name = "<span style=""color:#999999;"">Excel Workbook</span>"
                    Set tXlsxOpt.hover = $ZCVT(tXlsxReason,"O","HTML")
                }
                Do tOptions.%Push(tXlsxOpt)

                Set tJsonOpt = {"name":"JSON"}
                Set tJsonOpt.click = "zenPage.getComponent("_..index_").exportJSON();"
                Set tJsonOpt.hover = "JSON (.json)"
                Do tOptions.%Push(tJsonOpt)

                Set tHelp = $$$Text("Export current results","%DeepSee")
                Set tHtml = ##class(%ZEN.Component.html).%New()
                Set tHtml.content = ..%CreateExportDropdown(..index, 1, "deepsee/ds2_action_44.png", tHelp, tOptions)

                Do pGroup.%AddChildBefore(tHtml, tImgChild)
                Do pGroup.%RemoveChild(tImgChild)
            }
        }
    }
    Catch ex {
        // Non-fatal - see method doc comment.
    }

    Quit tSC
}

/// Fixes https://github.com/asinay/bi-export-plus/issues/1: the Dashboard Editor's
/// widget Settings panel resolves a widget's property schema via
/// <method>%DeepSee.Dashboard.Utils::%GetWidgetClass</method>, which matches the
/// dashboard widget's own <property>type</property> value against the
/// <var>pInfo</var> catalog <method>%GetCatalogInfo</method> builds for every
/// subclass of <class>%DeepSee.Component.Widget.widget</class> (confirmed live by
/// dumping <class>%DeepSee.Dashboard.Utils</class>' and
/// <class>%DeepSee.UserPortal.DashboardViewer</class>'s compiled source) - but the
/// stock <method>%GetCatalogInfo</method> this class inherits registers that catalog
/// entry under <method>%ClassName(0)</method> (the short, package-stripped class name,
/// "pivot" for both the stock class and this one). <class>BIExport.DashboardTools</class>
/// sets a converted widget's <property>type</property> to this class's *full* name
/// ("BIExport.UI.Widget.pivot" - needed because the dashboard *viewer*'s
/// <method>%CreateVisualWidget</method> uses <property>type</property> directly as a
/// class reference, not a catalog keyword), which never equals the inherited "pivot"
/// catalog entry - so <method>%GetWidgetClass</method> finds no match, the Editor's
/// settings lookup resolves to an empty class, and the Settings panel renders with no
/// properties at all. Registering this class's own catalog entry under its full name
/// instead - the exact string <method>ConvertWidget</method> writes into
/// <property>type</property> - lets <method>%GetWidgetClass</method> find it;
/// <method>%GetWidgetPropertyInfo</method> is left inherited unchanged, since the
/// property list itself is identical to the stock pivot widget's.
ClassMethod %GetCatalogInfo(Output pInfo As %List, Output pSubtypeList As %List) As %Status
{
    Set tSC = ##super(.pInfo, .pSubtypeList)
    Set pInfo = $LB(..%ClassName(1), $LG(pInfo,2), $LG(pInfo,3), $LG(pInfo,4), $LG(pInfo,5))
    Quit tSC
}

/// Runs after the stock widget finishes its own load logic; DOM-patches the existing
/// "csv"=1 Export dropdown to add "Excel Workbook"/"JSON", exactly like
/// <class>BIExport.UI.Analyzer</class>'s own <method>addExtraExportOptions</method> for
/// the full-page Analyzer. Two differences from that version, both because this runs
/// once per widget on a dashboard that can hold several:
///   - Scoped to this widget's own DOM subtree (<method>getEnclosingDiv</method>) rather
///     than the whole document, so multiple pivot widgets on one dashboard don't step on
///     each other's dropdowns.
///   - Also has to recognize (and skip, marking it so it's never rechecked) a dropdown
///     that already has "Excel Workbook"/JSON baked in server-side by this class's own
///     <method>%OnCreateControls</method> override above (the single-icon-widget case) -
///     otherwise this client-side patch, meant only for genuine csv=1 widgets that still
///     render the unmodified stock dropdown, would double-insert into that one.
ClientMethod onloadHandler() [ Language = javascript ]
{
    this.invokeSuper('onloadHandler', arguments);
    this.addExtraExportOptions();
}

/// Same synchronous round-trip as <class>BIExport.UI.Analyzer</class>'s own method of this
/// name - see that class's doc comment for how this was confirmed live (a
/// <parameter>ZenMethod</parameter> called with no callback argument returns its value
/// synchronously in this IRIS version's Zen). Defined here too, rather than only on the
/// page, because a widget's own client-side stub set only includes methods declared on the
/// widget's OWN class (or an ancestor) - confirmed live the same way, by calling this from
/// <method>addExtraExportOptions</method> below as <code>this.IsXLSXAvailable()</code> (the
/// widget instance), not <code>zenPage.IsXLSXAvailable()</code> (the page).
ClassMethod IsXLSXAvailable() As %String [ ZenMethod ]
{
    Set tOK = ##class(BIExport.XLSXWriter).IsAvailable(.tReason)
    Quit $Select(tOK:"1",1:"0|"_tReason)
}

ClientMethod addExtraExportOptions() [ Language = javascript ]
{
    try {
        var container = this.getEnclosingDiv();
        if (!container) { return; }

        // One round-trip per widget instance (not once per menu found below - there's only
        // ever one export dropdown per widget anyway).
        var xlsxStatus = this.IsXLSXAvailable();
        var xlsxAvailable = (xlsxStatus.charAt(0) === '1');
        var xlsxReason = xlsxAvailable ? '' : xlsxStatus.substring(2);

        var menus = container.querySelectorAll('.dsptIconDropdownMenu');
        for (var i = 0; i < menus.length; i++) {
            var menu = menus[i];
            if (menu.getAttribute('data-extraExportAdded')) { continue; }

            var links = menu.getElementsByTagName('a');
            var excelLink = null;
            var alreadyHasXLSX = false;
            for (var j = 0; j < links.length; j++) {
                var onclickAttr = links[j].getAttribute('onclick') || '';
                if (!excelLink && onclickAttr.indexOf('exportExcel') >= 0) {
                    excelLink = links[j];
                }
                if (onclickAttr.indexOf('exportXLSX') >= 0) {
                    alreadyHasXLSX = true;
                }
            }

            // Server-rendered by %OnCreateControls above (single-icon-widget case) -
            // already has our options, just mark it so it's skipped on future passes.
            if (alreadyHasXLSX) {
                menu.setAttribute('data-extraExportAdded', '1');
                continue;
            }
            if (!excelLink) { continue; }

            var match = (excelLink.getAttribute('onclick') || '').match(/toggleExportMenu\('([^']+)'\)/);
            var menuId = match ? match[1] : '';
            var prefix = 'zenPage.getComponent(' + this.index + ').toggleExportMenu(\'' + menuId + '\');zenPage.getComponent(' + this.index + ').';

            var xlsxLink = document.createElement('a');
            if (xlsxAvailable) {
                xlsxLink.setAttribute('title', 'Microsoft Excel Workbook (.xlsx)');
                xlsxLink.setAttribute('onclick', prefix + 'exportXLSX();');
            } else {
                xlsxLink.setAttribute('title', xlsxReason);
                xlsxLink.style.color = '#999999';
                xlsxLink.style.cursor = 'not-allowed';
            }
            xlsxLink.textContent = 'Excel Workbook';
            excelLink.parentNode.insertBefore(xlsxLink, excelLink.nextSibling);

            var jsonLink = document.createElement('a');
            jsonLink.setAttribute('title', 'JSON (.json)');
            jsonLink.setAttribute('onclick', prefix + 'exportJSON();');
            jsonLink.textContent = 'JSON';
            xlsxLink.parentNode.insertBefore(jsonLink, xlsxLink.nextSibling);

            menu.setAttribute('data-extraExportAdded', '1');
        }
    }
    catch (e) {
        // Non-fatal: if the stock widget's DOM structure ever changes enough that this
        // stops finding the menu, just skip adding the options rather than break the widget.
    }
}

/// Opens the real .xlsx export page for this widget's current pivot. See
/// buildExportURL for why this opens a same-application CSP page instead of calling
/// BIExport.REST.
ClientMethod exportXLSX() [ Language = javascript ]
{
    var url = this.buildExportURL('BIExport.UI.XLSXPage.cls', 'xlsx');
    if (url) { window.open(url, 'xlsx', ''); }
}

/// Opens the JSON export page for this widget's current pivot. See buildExportURL for
/// why this opens a same-application CSP page instead of calling BIExport.REST.
ClientMethod exportJSON() [ Language = javascript ]
{
    var url = this.buildExportURL('BIExport.UI.JSONPage.cls', 'json');
    if (url) { window.open(url, 'json', ''); }
}

/// Widget-context port of <class>BIExport.UI.Analyzer</class>'s own
/// <method>buildExportURL</method>, ported line-for-line from that method and the
/// relevant (non-KPI, cubeName/queryKey) branch of the stock widget's own client-side
/// <method>exportExcel</method> (confirmed live by dumping its compiled source) for
/// anything specific to running inside a dashboard widget rather than the Analyzer page:
///   - <var>table</var> comes from <method>getPivot</method> (== <code>this.findComponent('table')</code>,
///     confirmed live), not <method>getPivot</method> on a full-page Analyzer.
///   - <var>title</var> falls back to the widget definition's own title
///     (<code>this.getDefinition().title</code>) before <var>table.printTitle</var> -
///     exactly the stock exportExcel()'s own priority order - since a dashboard widget
///     commonly has its own title distinct from the underlying saved pivot's.
///   - Nonce creation calls <code>zenPage.CreateQueryNonce</code> (a page-level method),
///     not <code>this.CreateQueryNonce</code> - confirmed by the stock exportExcel()
///     doing the same, because <var>this</var> here is the widget, not the page.
/// Everything else (CUBE/LISTINGNAME, the crossed ROWTOTALS/COLUMNTOTALS/*AGG mapping -
/// see CLAUDE.md's widget-property-naming gotcha, ROWCAPTIONS, FILTERNAMES/FILTERVALUES,
/// FILENAME fallback) is identical to the Analyzer version.
ClientMethod buildExportURL(pPage, pExtension) [ Language = javascript ]
{
    var table = this.getPivot();
    if (!table || ('' == table.cubeName) || ('' == table.queryKey)) {
        return '';
    }

    var mdx = table.GetCurrentQueryText("export");
    var nonce = '';
    if (mdx.toString().length > 1000) {
        // avoid long urls for IE!
        nonce = zenPage.CreateQueryNonce(mdx);
        if (nonce.toString().indexOf('ERROR:') >= 0) {
            alert(nonce);
            return '';
        }
    }

    var url;
    if (nonce) {
        url = pPage + '?NONCE=' + encodeURIComponent(nonce);
    }
    else {
        url = pPage + '?MDX=' + encodeURIComponent(mdx);
    }

    var widgetDef = this.getDefinition();
    var title = (widgetDef && widgetDef.title) ? widgetDef.title : table.printTitle;
    var subtitle = table.printSubtitle;
    // Strip any BI-folder path prefix ("Test/Crossjoin...") - same reasoning as
    // BIExport.UI.Analyzer.buildExportURL(): that's portal organization, not part of
    // the report's identity, and shouldn't leak into the filename or Info sheet title.
    if (title && title.indexOf('/') >= 0) {
        var pathParts = title.split('/');
        title = pathParts[pathParts.length - 1];
    }
    url += '&TITLE=' + encodeURIComponent(title || '');
    url += '&SUBTITLE=' + encodeURIComponent(subtitle || '');
    url += '&CUBE=' + encodeURIComponent(table.cubeName || '');
    if (table.listing != '') {
        url += '&LISTINGNAME=' + encodeURIComponent(table.listing);
    }
    // Crossed ROWTOTALS/COLUMNTOTALS(+AGG) mapping - see CLAUDE.md gotcha #8 and
    // BIExport.UI.Analyzer.buildExportURL()'s own comment for the full explanation of
    // why table.rowTotals/table.columnTotals are named the opposite of what they render.
    if (table.rowTotals) {
        url += '&COLUMNTOTALS=1';
        if (table.rowTotalAgg) {
            url += '&COLUMNTOTALAGG=' + encodeURIComponent(table.rowTotalAgg);
        }
    }
    if (table.columnTotals) {
        url += '&ROWTOTALS=1';
        if (table.columnTotalAgg) {
            url += '&ROWTOTALAGG=' + encodeURIComponent(table.columnTotalAgg);
        }
    }
    url += '&ROWCAPTIONS=' + encodeURIComponent(table.getRowCaptions().join('\n'));

    var filterNames = [];
    var filterValues = [];
    table.getFilterInfo(filterNames, filterValues);
    if (filterNames.length > 0) {
        url += '&FILTERNAMES=' + encodeURIComponent(filterNames.join('\n'));
        url += '&FILTERVALUES=' + encodeURIComponent(filterValues.join('\n'));
    }

    var fileBase;
    if (title) {
        fileBase = title;
    }
    else {
        // No title anywhere (ad-hoc widget pivot, no widget title set) - fall back to
        // the cube plus a short timestamp instead of a bare "export.xlsx" that every
        // such download would collide on. Same fallback as the Analyzer version.
        var now = new Date();
        var pad = function(n) { return (n < 10 ? '0' : '') + n; };
        var ts = now.getFullYear() + pad(now.getMonth() + 1) + pad(now.getDate()) + '-' + pad(now.getHours()) + pad(now.getMinutes());
        fileBase = (table.cubeName ? table.cubeName : 'export') + '_' + ts;
    }
    url += '&FILENAME=' + encodeURIComponent(fileBase + '.' + pExtension);

    return zenLink(url);
}

}
