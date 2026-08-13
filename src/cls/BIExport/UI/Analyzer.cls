/// Thin subclass of the stock BI Analyzer page that adds real "Excel Workbook" (.xlsx)
/// and "JSON" options to the existing Export dropdown (alongside Excel/CSV/CSV with
/// Heading), calling <class>BIExport.UI.XLSXPage</class>/<class>BIExport.UI.JSONPage</class>
/// instead of the HTML-saved-as-.xls exporter.
///
/// Reachable at the same CSP application as the stock page, just under this class's
/// name instead of the "_DeepSee.UI.Analyzer.zen" alias, e.g.
/// /csp/<namespace>/BIExport.UI.Analyzer.cls?CUBE=... - point any links/shortcuts at
/// this URL instead of the stock one to pick up the new options.
Class BIExport.UI.Analyzer Extends %DeepSee.UI.Analyzer
{

/// Runs after the stock page finishes its own load logic; DOM-patches the existing
/// Export dropdown to add "Excel Workbook" and "JSON" items. Implemented as a DOM
/// patch (rather than duplicating the dropdown-building code) so it keeps working
/// even if the stock page's internal markup/ids change across versions.
ClientMethod onloadHandler() [ Language = javascript ]
{
    this.invokeSuper('onloadHandler', arguments);
    this.addExtraExportOptions();
}

ClientMethod addExtraExportOptions() [ Language = javascript ]
{
    try {
        var menus = document.querySelectorAll('.dsptIconDropdownMenu');
        for (var i = 0; i < menus.length; i++) {
            var menu = menus[i];
            if (menu.getAttribute('data-extraExportAdded')) { continue; }

            var links = menu.getElementsByTagName('a');
            var excelLink = null;
            for (var j = 0; j < links.length; j++) {
                var onclickAttr = links[j].getAttribute('onclick') || '';
                if (onclickAttr.indexOf('exportExcel') >= 0) {
                    excelLink = links[j];
                    break;
                }
            }
            if (!excelLink) { continue; }

            var match = (excelLink.getAttribute('onclick') || '').match(/toggleExportMenu\('([^']+)'\)/);
            var menuId = match ? match[1] : '';

            var xlsxLink = document.createElement('a');
            xlsxLink.setAttribute('title', 'Microsoft Excel Workbook (.xlsx)');
            xlsxLink.setAttribute('onclick', "zenPage.toggleExportMenu('" + menuId + "');zenPage.exportXLSX();");
            xlsxLink.textContent = 'Excel Workbook';
            excelLink.parentNode.insertBefore(xlsxLink, excelLink.nextSibling);

            var jsonLink = document.createElement('a');
            jsonLink.setAttribute('title', 'JSON (.json)');
            jsonLink.setAttribute('onclick', "zenPage.toggleExportMenu('" + menuId + "');zenPage.exportJSON();");
            jsonLink.textContent = 'JSON';
            xlsxLink.parentNode.insertBefore(jsonLink, xlsxLink.nextSibling);

            menu.setAttribute('data-extraExportAdded', '1');
        }
    }
    catch (e) {
        // Non-fatal: if the stock page's DOM structure ever changes enough that this
        // stops finding the menu, just skip adding the options rather than break the page.
    }
}

/// Opens the real .xlsx export page for the current pivot. See buildExportURL for
/// why this opens a same-application CSP page instead of calling BIExport.REST.
ClientMethod exportXLSX() [ Language = javascript ]
{
    var url = this.buildExportURL('BIExport.UI.XLSXPage.cls', 'xlsx');
    if (url) { window.open(url, 'xlsx', ''); }
}

/// Opens the JSON export page for the current pivot. See buildExportURL for why
/// this opens a same-application CSP page instead of calling BIExport.REST.
ClientMethod exportJSON() [ Language = javascript ]
{
    var url = this.buildExportURL('BIExport.UI.JSONPage.cls', 'json');
    if (url) { window.open(url, 'json', ''); }
}

/// Mirrors exportExcel()'s gathering of the current pivot's MDX/title, but points at
/// one of our own plain CSP pages (in this SAME application, not the separate
/// BIExport.REST application) so the browser's existing Analyzer login session covers
/// it - exactly the same reason exportExcel() itself uses window.open() to a sibling
/// page rather than a cross-application fetch. Shared by exportXLSX()/exportJSON();
/// pPage is the target page class name, pExtension the default filename extension.
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
        nonce = this.CreateQueryNonce(mdx);
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

    var title = table.printTitle;
    var subtitle = table.printSubtitle;
    if (title == '') {
        title = this.pivotName;
        if (title != '') {
            var s = title.toString().split('.');
            s.length = s.length - 1;
            title = s.join('.');
        }
    }
    // Strip any BI-folder path prefix ("Test/Crossjoin...") - that's portal
    // organization, not part of the report's identity, and shouldn't leak into
    // the filename or the Info sheet's title.
    if (title.indexOf('/') >= 0) {
        var pathParts = title.split('/');
        title = pathParts[pathParts.length - 1];
    }
    url += '&TITLE=' + encodeURIComponent(title);
    url += '&SUBTITLE=' + encodeURIComponent(subtitle);
    url += '&CUBE=' + encodeURIComponent(table.cubeName || '');
    if (table.listing != '') {
        url += '&LISTINGNAME=' + encodeURIComponent(table.listing);
    }
    if (table.rowTotals || table.columnTotals) {
        url += '&TOTALS=1';
    }
    // Row axis dimension/level captions (e.g. "YearSold","Product Category") - the
    // ResultSet API has no way to derive these; the pivot table widget already has
    // them client-side. Same source stock exportExcel() uses for its ROWCAPTIONS param.
    url += '&ROWCAPTIONS=' + encodeURIComponent(table.getRowCaptions().join('\n'));

    // The pivot's active filter selections (e.g. Product Category = Cereal+Dairy+...)
    // for the Info sheet's recap block - same table.getFilterInfo() call stock
    // exportExcel() uses for its FILTERNAMES/FILTERVALUES params.
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
        // Ad-hoc pivot built in the Analyzer and never saved - no name to hang a
        // filename on. Fall back to the cube plus a short timestamp instead of a
        // bare "export.xlsx" that every such download would collide on.
        var now = new Date();
        var pad = function(n) { return (n < 10 ? '0' : '') + n; };
        var ts = now.getFullYear() + pad(now.getMonth() + 1) + pad(now.getDate()) + '-' + pad(now.getHours()) + pad(now.getMinutes());
        fileBase = (table.cubeName ? table.cubeName : 'export') + '_' + ts;
    }
    url += '&FILENAME=' + encodeURIComponent(fileBase + '.' + pExtension);

    return zenLink(url);
}

}
