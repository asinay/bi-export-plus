/// A standalone demo page for showing off <class>BIExport.XLSXWriter</class>/<class>
/// BIExport.MDXExport</class> without needing to build a pivot in the Analyzer first.
/// Pick a canned sample query (or paste your own MDX), toggle totals/aggregation, and
/// download a real .xlsx or .json - same code path as the Analyzer's own "Excel
/// Workbook"/"JSON" export buttons (<class>BIExport.UI.XLSXPage</class>/<class>
/// BIExport.UI.JSONPage</class>), just triggered from a plain form instead of a live
/// pivot widget. Lives in the same CSP application as the Analyzer, so it shares the
/// same login session - no separate REST auth needed for the demo.
///
/// Not part of the feature itself - a demo/sales-tool convenience only.
Class BIExport.UI.DemoPage Extends %CSP.Page
{

Parameter CONTENTTYPE = "text/html";

ClassMethod OnPage() As %Status
{
    Set tXLSXAvailable = ##class(BIExport.XLSXWriter).IsAvailable(.tReason)

    &html<<!DOCTYPE html>
<html>
<head>
<title>bi-export-plus - Live Demo</title>
<meta charset="utf-8">
<style>
  body { font-family: Calibri, "Segoe UI", Arial, sans-serif; margin: 0; background: #F4F6FA; color: #1B2A4A; }
  header { background: #1B2A4A; color: #fff; padding: 18px 32px; }
  header h1 { margin: 0; font-size: 22px; }
  header p { margin: 4px 0 0; color: #AAB8D6; font-size: 13px; }
  main { max-width: 780px; margin: 28px auto; padding: 0 20px; }
  .card { background: #fff; border-radius: 8px; padding: 24px 28px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); margin-bottom: 20px; }
  .card h2 { font-size: 15px; margin: 0 0 14px; color: #1B2A4A; }
  label { display: block; font-size: 13px; font-weight: bold; margin: 12px 0 4px; }
  select, textarea, input[type=text] { width: 100%; box-sizing: border-box; font-family: Consolas, monospace; font-size: 13px; padding: 8px; border: 1px solid #D9D9D9; border-radius: 4px; }
  textarea { height: 90px; resize: vertical; }
  .totals-row { display: flex; align-items: center; gap: 10px; margin: 10px 0; }
  .totals-row input[type=checkbox] { width: 16px; height: 16px; }
  .totals-row select { width: auto; flex: 1; }
  .totals-row .lbl { font-size: 13px; font-weight: normal; flex: 2; }
  .buttons { margin-top: 18px; display: flex; gap: 12px; }
  button { flex: 1; padding: 12px; font-size: 14px; font-weight: bold; border: none; border-radius: 5px; cursor: pointer; color: #fff; }
  button:disabled { background: #B0B7C3; cursor: not-allowed; }
  #btnXlsx { background: #4472C4; }
  #btnXlsx:hover { background: #35599e; }
  #btnJson { background: #595959; }
  #btnJson:hover { background: #404040; }
  .hint { font-size: 12px; color: #808080; margin-top: 6px; }
  .banner { background: #FDEDED; border: 1px solid #F5C2C2; border-radius: 6px; padding: 12px 16px; margin-bottom: 20px; font-size: 13px; }
  .banner b { display: block; margin-bottom: 4px; }
</style>
</head>
<body>
<header>
  <h1>bi-export-plus</h1>
  <p>Live demo &mdash; real .xlsx / .json exports straight from MDX, no pivot required</p>
</header>
<main>>

    If ('tXLSXAvailable) {
        &html<<div class="banner">
    <b>Excel Workbook export is disabled - see below. JSON export still works.</b>
    #($ZCVT(tReason,"O","HTML"))#
  </div>>
    }

    &html<<div class="card">
    <h2>1. Pick a query</h2>
    <select id="preset" onchange="applyPreset()">
      <option value="simple">Simple &mdash; Revenue &amp; Units by Region</option>
      <option value="totals">With totals &mdash; Sales by Year</option>
      <option value="nested">Nested column headers &mdash; Sales by Year and Channel</option>
      <option value="custom">Custom MDX&hellip;</option>
    </select>

    <label for="mdx">MDX query</label>
    <textarea id="mdx" oninput="syncFiltersFromMdx()"></textarea>

    <label for="title">Report title</label>
    <input type="text" id="title">

    <label for="filters">Filters (one per line, <code>Name = Value</code>)</label>
    <textarea id="filters" placeholder="Region = West"></textarea>
    <p class="hint">Auto-parsed from the query's <code>%FILTER</code> clause when you paste MDX (e.g. straight from the Analyzer) - shows raw member keys, not resolved friendly captions, since that's all a bare MDX string carries. Still shown as an editable recap block on the exported Info sheet, so tweak it by hand if you want nicer labels/values.</p>
  </div>

  <div class="card">
    <h2>2. Totals</h2>
    <div class="totals-row">
      <input type="checkbox" id="rowTotals">
      <span class="lbl">Total <b>column</b> (one value per row, summed across columns)</span>
      <select id="rowTotalAgg">
        <option value="sum">Sum</option>
        <option value="count">Count</option>
        <option value="min">Min</option>
        <option value="max">Max</option>
        <option value="avg">Average</option>
        <option value="pct">% of Total</option>
      </select>
    </div>
    <div class="totals-row">
      <input type="checkbox" id="columnTotals">
      <span class="lbl">Total <b>row</b> (one value per column, summed down rows)</span>
      <select id="columnTotalAgg">
        <option value="sum">Sum</option>
        <option value="count">Count</option>
        <option value="min">Min</option>
        <option value="max">Max</option>
        <option value="avg">Average</option>
        <option value="pct">% of Total</option>
      </select>
    </div>
    <p class="hint">Independent switches - turn on either, both, or neither.</p>
  </div>

  <div class="buttons">>

    If (tXLSXAvailable) {
        Write "<button id=""btnXlsx"" onclick=""doExport('xlsx')"">Download Excel Workbook</button>",!
    } Else {
        Write "<button id=""btnXlsx"" disabled title=""" _ $ZCVT(tReason,"O","HTML") _ """>Download Excel Workbook</button>",!
    }

    &html<<button id="btnJson" onclick="doExport('json')">Download JSON</button>
  </div>
</main>

<script>
var PRESETS = {
  simple: {
    mdx: "SELECT {[Measures].[Amount Sold],[Measures].[Units Sold]} ON COLUMNS, NON EMPTY [Outlet].[H1].[Region].Members ON ROWS FROM [HOLEFOODS]",
    title: "HoleFoods by Region",
    cube: "HOLEFOODS",
    rowCaptions: "Region"
  },
  totals: {
    mdx: "SELECT {[Measures].[%COUNT],[Measures].[Amount Sold],[Measures].[Units Sold]} ON COLUMNS, NON EMPTY [DateOfSale].[Actual].[YearSold].Members ON ROWS FROM [HOLEFOODS]",
    title: "Sales by Year",
    cube: "HOLEFOODS",
    rowCaptions: "YearSold"
  },
  nested: {
    mdx: "SELECT Crossjoin([Channel].Members,{[Measures].[%COUNT],[Measures].[Amount Sold],[Measures].[Units Sold]}) ON COLUMNS, NON EMPTY [DateOfSale].[Actual].[YearSold].Members ON ROWS FROM [HOLEFOODS]",
    title: "Sales by Year and Channel",
    cube: "HOLEFOODS",
    rowCaptions: "YearSold"
  },
  custom: { mdx: "", title: "", cube: "HOLEFOODS", rowCaptions: "" }
};

// Best-effort: pulls member keys out of the query's %FILTER clause by regex, grouped
// by their innermost level name (the bracket segment right before the final ".&[key]"
// or ".[MemberName]"). Raw MDX carries no friendly dimension/level captions - the
// %DeepSee.ResultSet API can't derive those either (see CLAUDE.md) - so this shows
// keys as they appear in the query, not resolved display names.
function extractFiltersFromMDX(mdx) {
  var idx = mdx.search(/%FILTER\b/i);
  if (idx === -1) { return []; }
  var clause = mdx.slice(idx);
  var re = /((?:\[[^\]]+\]\.)+)&?\[([^\]]+)\]/g;
  var order = [];
  var groups = {};
  var m;
  while ((m = re.exec(clause)) !== null) {
    var segments = m[1].match(/\[([^\]]+)\]/g).map(function (s) { return s.slice(1, -1); });
    var level = segments[segments.length - 1];
    var key = m[2];
    if (!groups[level]) { groups[level] = []; order.push(level); }
    groups[level].push(key);
  }
  return order.map(function (level) { return level + ' = ' + groups[level].join(', '); });
}

function syncFiltersFromMdx() {
  document.getElementById('filters').value = extractFiltersFromMDX(document.getElementById('mdx').value).join('\n');
}

function applyPreset() {
  var p = PRESETS[document.getElementById('preset').value];
  document.getElementById('mdx').value = p.mdx;
  document.getElementById('title').value = p.title;
  syncFiltersFromMdx();
}

function doExport(ext) {
  var mdx = document.getElementById('mdx').value;
  if (!mdx) { alert('Enter an MDX query first.'); return; }
  var preset = PRESETS[document.getElementById('preset').value] || PRESETS.custom;
  var title = document.getElementById('title').value;

  var page = (ext === 'xlsx') ? 'BIExport.UI.XLSXPage.cls' : 'BIExport.UI.JSONPage.cls';
  var url = page + '?MDX=' + encodeURIComponent(mdx);
  url += '&TITLE=' + encodeURIComponent(title);
  url += '&CUBE=' + encodeURIComponent(preset.cube || '');
  if (preset.rowCaptions) {
    url += '&ROWCAPTIONS=' + encodeURIComponent(preset.rowCaptions);
  }

  var filterNames = [];
  var filterValues = [];
  document.getElementById('filters').value.split('\n').forEach(function (line) {
    var eq = line.indexOf('=');
    if (eq === -1) { return; }
    var name = line.slice(0, eq).trim();
    var value = line.slice(eq + 1).trim();
    if (!name) { return; }
    filterNames.push(name);
    filterValues.push(value);
  });
  if (filterNames.length > 0) {
    url += '&FILTERNAMES=' + encodeURIComponent(filterNames.join('\n'));
    url += '&FILTERVALUES=' + encodeURIComponent(filterValues.join('\n'));
  }

  if (document.getElementById('rowTotals').checked) {
    url += '&ROWTOTALS=1&ROWTOTALAGG=' + document.getElementById('rowTotalAgg').value;
  }
  if (document.getElementById('columnTotals').checked) {
    url += '&COLUMNTOTALS=1&COLUMNTOTALAGG=' + document.getElementById('columnTotalAgg').value;
  }
  url += '&FILENAME=' + encodeURIComponent((title || 'export') + '.' + ext);
  window.open(url, '_blank');
}

applyPreset();
</script>
</body>
</html>>
    Quit $$$OK
}

}
