/// Renders the normalized JSON structure produced by <class>BIExport.MDXExport</class>
/// into a genuine .xlsx workbook (real cell types and number formats - not an HTML
/// table saved with an .xls extension).
///
/// Layout: the data sheet is a real Excel structured Table (native autofilter/banded
/// rows, sortable, recognized as a table by Excel/other tools) starting at A1 - the
/// header row is never pushed down by metadata. Title/subtitle/filters instead go on
/// a separate "Info" sheet, plus the workbook's own document properties.
///
/// <method>BuildXLSX</method> (single pivot/listing) and <method>BuildMultiSheetXLSX</method>
/// (one workbook per dashboard, one sheet per pivot widget - see
/// <class>BIExport.DashboardExport</class>) share the actual per-sheet writing logic via
/// <method>%WriteResultSheet</method>, a private <parameter>Language = python</parameter>
/// classmethod, rather than duplicating it - confirmed live, before writing either
/// caller, that a live <class>openpyxl</class> <code>Worksheet</code> object built inside
/// one <parameter>Language = python</parameter> classmethod can be passed as a plain
/// argument into another <parameter>Language = python</parameter> classmethod of this
/// same class (via <code>iris.cls('BIExport.XLSXWriter')._WriteResultSheet(ws, data, table_name)</code> -
/// note the leading <code>%</code> becomes a leading <code>_</code> for Python attribute
/// access, confirmed the same way) and mutated there exactly as if the call had never
/// left Python - no serialization,
/// no loss of identity, <parameter>Private</parameter> classmethods included. That
/// confirmation is what made the shared-helper approach (rather than duplicating this
/// logic per the plan's own fallback) safe to build directly.
Class BIExport.XLSXWriter Extends %RegisteredObject
{

/// Two-tier check for whether <method>BuildXLSX</method>/<method>BuildMultiSheetXLSX</method>
/// can actually run: plain ObjectScript (no <parameter>Language = python</parameter>), so it's
/// safe to call from anywhere - a server-side gray-out check, a REST/CSP page's top-of-method
/// guard, or a client-side ZenMethod round-trip - without dragging Python into the call chain
/// just to answer "is Python even there." Confirmed live (see docs/iris-environment.md and
/// CLAUDE.md's Embedded Python section) that <code>##class(%SYS.Python).Import(...)</code>
/// raises a catchable <class>%Exception.PythonException</class> on failure (e.g.
/// <code>ModuleNotFoundError</code> for a missing package) rather than returning a %Status or
/// writing to some other error channel - this method leans on that directly.
///
/// Two distinct failure modes, each needing a different fix, so <var>pReason</var> is worded
/// accordingly:
///   - <code>Import("sys")</code> fails: Embedded Python itself isn't configured on this
///     instance (the CPF <code>PythonRuntimeLibrary</code>/<code>PythonRuntimeLibraryVersion</code>
///     parameters are empty or wrong) - a restart is required to pick up a CPF edit, so
///     <var>pReason</var> says so rather than suggesting anything short of that.
///   - <code>Import("sys")</code> succeeds but <code>Import("openpyxl")</code> fails: Python
///     itself is fine, the package just isn't installed for whichever Python
///     <code>PythonRuntimeLibrary</code> points at - <var>pReason</var> says to
///     <code>pip install openpyxl</code> against that same Python.
///
/// No caching - availability only changes via a CPF edit + restart (an extremely rare event
/// relative to how often this gets called), and a fresh <code>Import()</code> on an
/// already-loaded module is expected to be cheap; add caching later only if actually measured
/// to matter.
ClassMethod IsAvailable(Output pReason As %String) As %Boolean
{
    Set pReason = ""
    Set tAvailable = 1
    Try {
        Do ##class(%SYS.Python).Import("sys")
    } Catch ex {
        Set pReason = "IRIS Embedded Python isn't configured on this instance ("_ex.DisplayString()_"). "
            _ "Set the CPF parameters PythonRuntimeLibrary (path to a supported Python's shared "
            _ "library, e.g. python313.dll) and PythonRuntimeLibraryVersion (major.minor only, "
            _ "e.g. ""3.13"") via Config.config, then restart the instance - see "
            _ "docs/iris-environment.md's ""Embedded Python - required a fix"" section."
        Set tAvailable = 0
    }
    If (tAvailable) {
        Try {
            Do ##class(%SYS.Python).Import("openpyxl")
        } Catch ex {
            Set pReason = "IRIS Embedded Python is configured, but the openpyxl package isn't "
                _ "importable from it ("_ex.DisplayString()_"). Run pip install openpyxl against "
                _ "whichever Python installation the PythonRuntimeLibrary CPF parameter points at."
            Set tAvailable = 0
        }
    }
    Quit tAvailable
}

/// pJSON is the %ToJSON() output of a BIExport.MDXExport.BuildExportData() result.
/// Returns the raw bytes of an .xlsx file.
ClassMethod BuildXLSX(pJSON As %String) As %Binary [ Language = python ]
{
import json
import io
import openpyxl
from openpyxl.styles import Font, PatternFill
import iris

data = json.loads(pJSON)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Export"

INFO_TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
INFO_TITLE_FILL = PatternFill("solid", fgColor="4472C4")
INFO_LABEL_FILL = PatternFill("solid", fgColor="D9E1F2")
INFO_NOTE_FONT = Font(italic=True, color="808080")
HEADER_FONT = Font(bold=True)

# ---------------------------------------------------------------------------
# Info sheet + document properties - title/subtitle/filters live here, never
# inline on the data sheet, so the data sheet's header always starts at A1.
# ---------------------------------------------------------------------------
title = data.get("title") or ""
subtitle = data.get("subtitle") or ""
cube = data.get("cube") or ""
filters = data.get("filters") or []

info = wb.create_sheet("Info")
info.sheet_view.showGridLines = False

info.cell(row=1, column=1, value="Export Info")
info.cell(row=1, column=2, value="")
info.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
for col in (1, 2):
    c = info.cell(row=1, column=col)
    c.font = INFO_TITLE_FONT
    c.fill = INFO_TITLE_FILL
info.row_dimensions[1].height = 22

def info_label(row, text):
    c = info.cell(row=row, column=1, value=text)
    c.font = HEADER_FONT
    c.fill = INFO_LABEL_FILL

r = 3
if title:
    info_label(r, "Report")
    info.cell(row=r, column=2, value=title)
    r += 1
if cube:
    info_label(r, "Cube")
    info.cell(row=r, column=2, value=cube)
    r += 1
if subtitle:
    info_label(r, "Subtitle")
    info.cell(row=r, column=2, value=subtitle)
    r += 1
for f in filters:
    info_label(r, f.get("name", ""))
    info.cell(row=r, column=2, value=f.get("value", ""))
    r += 1
if r == 3:
    c = info.cell(row=r, column=1, value="(no title, cube, subtitle, or filters supplied for this export)")
    c.font = INFO_NOTE_FONT
info.column_dimensions["A"].width = 24
info.column_dimensions["B"].width = 50

if title:
    wb.properties.title = title
if subtitle:
    wb.properties.description = subtitle
if cube:
    wb.properties.subject = cube

notes = iris.cls('BIExport.XLSXWriter')._WriteResultSheet(ws, data, "ExportTable")
for note in notes:
    note_cell = info.cell(row=info.max_row + 2, column=1, value=note)
    note_cell.font = INFO_NOTE_FONT

# Branding footer on the Info sheet - computed from info.max_row rather than a
# manually-tracked counter, since earlier code writes a variable, optional number of
# rows there (title/cube/subtitle/filters, plus any notes just above) and none of it
# updates a shared "next free row" after writing.
BRAND_LINK = "https://github.com/asinay/bi-export-plus"
brand_row = info.max_row + 2
brand_cell = info.cell(row=brand_row, column=1, value="Generated with bi-export-plus")
brand_cell.font = INFO_NOTE_FONT
link_cell = info.cell(row=brand_row + 1, column=1, value=BRAND_LINK)
link_cell.hyperlink = BRAND_LINK
link_cell.font = Font(color="0563C1", underline="single")

buf = io.BytesIO()
wb.save(buf)
return buf.getvalue()
}

/// pJSON is the %ToJSON() output of a BIExport.DashboardExport.BuildDashboardExportData()
/// result: {title, dashboardName, sheets:[{widgetName,title,data|error}], skipped:[...]}.
/// Builds one combined "Info" sheet (dashboard title, a recap row per widget - sheet
/// name, status, and any per-sheet note - plus the skipped-widgets list) and one data
/// sheet per widget that has a "data" entry, via the same <method>%WriteResultSheet</method>
/// helper <method>BuildXLSX</method> uses. A widget entry with an "error" instead of
/// "data" (see <class>BIExport.DashboardExport</class>) gets a real sheet too, but with
/// just that error message written on it rather than a Table/merged header - so the
/// workbook always has exactly one sheet per pivot-family widget, whether or not that
/// widget's own export succeeded, plus the Info sheet.
///
/// Returns the raw bytes of an .xlsx file.
ClassMethod BuildMultiSheetXLSX(pJSON As %String) As %Binary [ Language = python ]
{
import json
import io
import openpyxl
from openpyxl.styles import Font, PatternFill
import iris

data = json.loads(pJSON)

wb = openpyxl.Workbook()
info = wb.active
info.title = "Info"
info.sheet_view.showGridLines = False

INFO_TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
INFO_TITLE_FILL = PatternFill("solid", fgColor="4472C4")
INFO_LABEL_FILL = PatternFill("solid", fgColor="D9E1F2")
INFO_NOTE_FONT = Font(italic=True, color="808080")
HEADER_FONT = Font(bold=True)
ERROR_FONT = Font(color="C00000")

dashboard_title = data.get("title") or ""
dashboard_name = data.get("dashboardName") or ""
sheets = data.get("sheets") or []
skipped = data.get("skipped") or []

info.cell(row=1, column=1, value="Export Info")
info.cell(row=1, column=2, value="")
info.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
for col in (1, 2):
    c = info.cell(row=1, column=col)
    c.font = INFO_TITLE_FONT
    c.fill = INFO_TITLE_FILL
info.row_dimensions[1].height = 22

def info_label(row, text):
    c = info.cell(row=row, column=1, value=text)
    c.font = HEADER_FONT
    c.fill = INFO_LABEL_FILL

r = 3
info_label(r, "Dashboard")
info.cell(row=r, column=2, value=dashboard_title or dashboard_name)
r += 1
if dashboard_title and dashboard_name != dashboard_title:
    info_label(r, "Path")
    info.cell(row=r, column=2, value=dashboard_name)
    r += 1
r += 1

info_label(r, "Widget")
info.cell(row=r, column=2, value="Sheet")
c = info.cell(row=r, column=3, value="Status")
for col in (1, 2, 3):
    info.cell(row=r, column=col).font = HEADER_FONT
r += 1

# Excel sheet-name rules (<=31 chars, none of []:*?/\, unique within the workbook) are
# enforced by %SanitizeSheetName; used_names below is just this loop's own running list
# so widget N+1's dedup sees widget N's already-chosen name.
def sanitize_sheet_name(name, used):
    return iris.cls('BIExport.XLSXWriter')._SanitizeSheetName(name, used)

used_names = ["Info"]

for sheet_index, sheet_entry in enumerate(sheets, start=1):
    widget_name = sheet_entry.get("widgetName") or ""
    sheet_title = sheet_entry.get("title") or widget_name or "Sheet"
    sheet_name = sanitize_sheet_name(sheet_title, used_names)
    used_names.append(sheet_name)

    ws = wb.create_sheet(sheet_name)

    info.cell(row=r, column=1, value=widget_name)
    info.cell(row=r, column=2, value=sheet_name)

    if "error" in sheet_entry:
        err_cell = ws.cell(row=1, column=1, value="Error exporting this widget:")
        err_cell.font = HEADER_FONT
        ws.cell(row=2, column=1, value=str(sheet_entry["error"]))
        ws.column_dimensions["A"].width = 100
        status_cell = info.cell(row=r, column=3, value="Error - see sheet")
        status_cell.font = ERROR_FONT
    else:
        # Excel Table names must be unique across the WHOLE workbook (not just per
        # sheet) - "Table_<n>" per widget guarantees that plus Excel's identifier rules
        # (alphanumeric/underscore, can't start with a digit), unlike the sheet's own
        # (arbitrary, sanitized-for-different-rules) display name.
        table_name = "Table_%d" % sheet_index
        notes = iris.cls('BIExport.XLSXWriter')._WriteResultSheet(ws, sheet_entry.get("data") or {}, table_name)
        info.cell(row=r, column=3, value="OK")
        for note in notes:
            r += 1
            note_cell = info.cell(row=r, column=2, value=note)
            note_cell.font = INFO_NOTE_FONT
    r += 1

if not sheets:
    info.cell(row=r, column=1, value="(no pivot widgets found on this dashboard)")
    info.cell(row=r, column=1).font = INFO_NOTE_FONT
    r += 1

if skipped:
    r += 1
    note_cell = info.cell(row=r, column=1,
               value="Skipped (not a pivot-family widget - chart/KPI/text widgets have "
                     "no exportable pivot data of their own):")
    note_cell.font = INFO_NOTE_FONT
    r += 1
    for sk in skipped:
        info.cell(row=r, column=1, value=sk.get("widgetName", ""))
        info.cell(row=r, column=2, value=sk.get("title", "") or sk.get("type", ""))
        r += 1

info.column_dimensions["A"].width = 24
info.column_dimensions["B"].width = 40
info.column_dimensions["C"].width = 18

if dashboard_title:
    wb.properties.title = dashboard_title

BRAND_LINK = "https://github.com/asinay/bi-export-plus"
brand_row = info.max_row + 2
brand_cell = info.cell(row=brand_row, column=1, value="Generated with bi-export-plus")
brand_cell.font = INFO_NOTE_FONT
link_cell = info.cell(row=brand_row + 1, column=1, value=BRAND_LINK)
link_cell.hyperlink = BRAND_LINK
link_cell.font = Font(color="0563C1", underline="single")

buf = io.BytesIO()
wb.save(buf)
return buf.getvalue()
}

/// Shared per-sheet writer for one BuildExportData()-shaped `data` dict, used by both
/// <method>BuildXLSX</method> (writing to its one "Export" sheet) and
/// <method>BuildMultiSheetXLSX</method> (writing to one sheet per widget). Builds either
/// a real Excel Table (the common case) or - when the column axis has more than one
/// level (<var>colHeaderDepth</var> > 1) - a true merged multi-row header with no Table
/// at all, exactly as the original single-sheet BuildXLSX always has (see the "Nested
/// column headers" tradeoff in README.md). Returns a (possibly empty) list of plain-text
/// notes the caller should surface somewhere on its own Info sheet - currently just the
/// nested-header/no-Table explanation, when that branch is taken.
///
/// pWs is a live openpyxl Worksheet already created (and, other than its title, still
/// blank) by the caller; this method writes into it directly and returns without saving
/// or otherwise touching the workbook. pTableName is the Excel Table's displayName (only
/// used when a real Table gets built at all - see the colHeaderDepth>1 branch below) -
/// Excel requires Table names to be unique across the WHOLE workbook, not just per
/// sheet, so a multi-sheet caller must pass a distinct name per widget; a single-sheet
/// caller can just pass the fixed "ExportTable" name this always used before this method
/// was extracted out of BuildXLSX.
ClassMethod %WriteResultSheet(pWs, pData, pTableName) [ Language = python, Private ]
{
import datetime
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ws = pWs
data = pData
notes = []

HEADER_FONT = Font(bold=True)
col_widths = {}

def touch_width(col_idx, text):
    length = len(str(text)) if text is not None else 0
    if length > col_widths.get(col_idx, 0):
        col_widths[col_idx] = length

def write_header(row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    c.font = HEADER_FONT
    touch_width(col, text)
    return c

def apply_cell(r, c, cellobj):
    cell = ws.cell(row=r, column=c)
    if cellobj is None:
        return cell
    t = cellobj.get("t", "s")
    v = cellobj.get("v")
    fmt = cellobj.get("fmt")
    if t == "n":
        cell.value = v
        if fmt:
            cell.number_format = fmt
    elif t == "d":
        if v:
            try:
                cell.value = datetime.date.fromisoformat(v)
            except ValueError:
                cell.value = v
            cell.number_format = fmt or "mm/dd/yyyy"
        else:
            cell.value = ""
    elif t == "t":
        cell.value = (v or 0) / 86400.0
        cell.number_format = fmt or "[h]:mm:ss"
    else:
        cell.value = v
        cell.number_format = "@"
    touch_width(c, v)
    return cell

def finalize_table(min_col, max_col, last_row, table_name, headers):
    """Wraps A1:.. through (min_col,max_col,last_row) as a real Excel Table. `headers`
    is the list of already-written header strings for min_col..max_col, used verbatim
    as the Table's column names (Excel requires these to match the header row cells)."""
    ref = "%s1:%s%d" % (get_column_letter(min_col), get_column_letter(max_col), last_row)
    tbl = Table(displayName=table_name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
    tbl._initialise_columns()
    for col, name in zip(tbl.tableColumns, headers):
        col.name = str(name) if name else col.name
    ws.add_table(tbl)
    return tbl

def safe_name(name, fallback):
    return name if name else fallback

kind = data.get("kind", "cellset")

if kind == "listing":
    columns = data.get("columns", [])
    rows = data.get("rows", [])
    if not columns or not rows:
        ws.cell(row=1, column=1, value="No results")
    else:
        headers = [safe_name(c.get("name", ""), "Column%d" % (j + 1)) for j, c in enumerate(columns)]
        for j, h in enumerate(headers):
            write_header(1, j + 1, h)
        for i, data_row in enumerate(rows):
            for j, val in enumerate(data_row):
                ws.cell(row=2 + i, column=j + 1, value=val)
                touch_width(j + 1, val)
        last_row = 1 + len(rows)
        finalize_table(1, len(headers), last_row, pTableName or "ExportTable", headers)
        ws.freeze_panes = "A2"
else:
    col_header_depth = data.get("colHeaderDepth", 0)
    row_header_depth = data.get("rowHeaderDepth", 0)
    col_count = data.get("colCount", 0)
    row_count = data.get("rowCount", 0)
    column_headers = data.get("columnHeaders", [])  # outer level first, inner (measure) last
    row_headers = data.get("rowHeaders", [])
    row_header_names = data.get("rowHeaderNames", [])
    cells = data.get("cells", [])
    column_totals = data.get("columnTotals")
    row_totals = data.get("rowTotals")
    grand_total = data.get("grandTotal")
    # Excel's own totalsRowFunction vocabulary (ECMA-376): none/sum/min/max/average/
    # count/countNums/stdDev/var/custom. Our aggregation keys match 1:1 except "avg"
    # (-> "average") and "pct" (% of total has no equivalent - "custom" means "not one
    # of Excel's built-in functions", which is accurate: there's no live formula behind
    # it, just the precomputed value).
    TOTALS_ROW_FUNCTION = {"sum": "sum", "count": "count", "min": "min", "max": "max",
                           "avg": "average", "pct": "custom"}
    col_totals_row_function = TOTALS_ROW_FUNCTION.get(data.get("columnTotalAgg", "sum"), "custom")
    row_totals_row_function = TOTALS_ROW_FUNCTION.get(data.get("rowTotalAgg", "sum"), "custom")

    if col_count == 0 or row_count == 0:
        ws.cell(row=1, column=1, value="No results")
    elif col_header_depth > 1:
        # A dimension on Columns needs a true 2+ row merged header - Excel's structured
        # Table feature requires exactly one header row, so this shape can't also be a
        # Table (see the "Nested column headers" tradeoff in README.md). Rather than
        # flatten the levels into one row (losing the visual grouping) to force a Table,
        # this branch keeps the real merged header and gives up the Table instead:
        # native autofilter/banded rows/tool-recognition, in exchange for an accurate
        # visual structure. Row-header column titles span every header row (merged
        # vertically); each column-header level gets its own row, merged horizontally
        # across runs of columns that share the same value at that level and every
        # shallower level (so a run never straddles an outer group's boundary).
        has_row_totals = row_totals is not None
        has_col_totals = column_totals is not None
        last_col = row_header_depth + col_count + (1 if has_row_totals else 0)
        header_rows = col_header_depth
        first_data_row = header_rows + 1

        for k in range(row_header_depth):
            name = row_header_names[k] if k < len(row_header_names) else ""
            write_header(1, k + 1, safe_name(name, "Row %d" % (k + 1)))
            ws.merge_cells(start_row=1, start_column=k + 1, end_row=header_rows, end_column=k + 1)

        def levels_match(a, b, upto):
            return all(column_headers[lv][a] == column_headers[lv][b] for lv in range(upto + 1))

        for lvl in range(col_header_depth):
            row = lvl + 1
            mergeable = lvl < col_header_depth - 1
            j = 0
            while j < col_count:
                k = j
                if mergeable:
                    while k + 1 < col_count and levels_match(j, k + 1, lvl):
                        k += 1
                col = row_header_depth + 1 + j
                write_header(row, col, column_headers[lvl][j])
                if k > j:
                    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=row_header_depth + 1 + k)
                j = k + 1

        if has_row_totals:
            write_header(1, row_header_depth + 1 + col_count, "Total")
            ws.merge_cells(start_row=1, start_column=row_header_depth + 1 + col_count,
                            end_row=header_rows, end_column=row_header_depth + 1 + col_count)

        # Data rows.
        for i in range(row_count):
            r = first_data_row + i
            hdrs = row_headers[i] if i < len(row_headers) else []
            for k, label in enumerate(hdrs):
                ws.cell(row=r, column=k + 1, value=label)
                touch_width(k + 1, label)
            row_cells = cells[i] if i < len(cells) else []
            for j, cellobj in enumerate(row_cells):
                apply_cell(r, row_header_depth + 1 + j, cellobj)
            if has_row_totals:
                apply_cell(r, row_header_depth + 1 + col_count, row_totals[i])

        last_data_row = first_data_row - 1 + row_count

        if has_col_totals:
            totals_row_num = last_data_row + 1
            if row_header_depth > 0:
                ws.cell(row=totals_row_num, column=row_header_depth, value="Total")
            for j, cellobj in enumerate(column_totals):
                apply_cell(totals_row_num, row_header_depth + 1 + j, cellobj)
            if has_row_totals and grand_total is not None:
                apply_cell(totals_row_num, row_header_depth + 1 + col_count, grand_total)
            for c in range(1, last_col + 1):
                ws.cell(row=totals_row_num, column=c).font = HEADER_FONT

        ws.freeze_panes = "%s%d" % (get_column_letter(row_header_depth + 1), first_data_row)

        notes.append("'%s': this pivot has a dimension on Columns (multi-level column "
                     "header), so its sheet keeps a true merged header instead of a "
                     "flattened Excel Table - no autofilter/banded rows on this sheet." % ws.title)
    else:
        has_row_totals = row_totals is not None
        has_col_totals = column_totals is not None
        total_col_count = col_count + (1 if has_row_totals else 0)
        last_col = row_header_depth + total_col_count

        # Header row (always row 1 - this is the whole point of the rewrite).
        row_headers_text = []
        for k in range(row_header_depth):
            name = row_header_names[k] if k < len(row_header_names) else ""
            name = safe_name(name, "Row %d" % (k + 1))
            row_headers_text.append(name)
            write_header(1, k + 1, name)

        col_headers_text = []
        for j in range(col_count):
            label = column_headers[0][j] if column_headers else ""
            col_headers_text.append(label)
            write_header(1, row_header_depth + 1 + j, label)

        if has_row_totals:
            write_header(1, row_header_depth + 1 + col_count, "Total")

        # Data rows.
        for i in range(row_count):
            r = 2 + i
            hdrs = row_headers[i] if i < len(row_headers) else []
            for k, label in enumerate(hdrs):
                ws.cell(row=r, column=k + 1, value=label)
                touch_width(k + 1, label)
            row_cells = cells[i] if i < len(cells) else []
            for j, cellobj in enumerate(row_cells):
                apply_cell(r, row_header_depth + 1 + j, cellobj)
            if has_row_totals:
                apply_cell(r, row_header_depth + 1 + col_count, row_totals[i])

        last_data_row = 1 + row_count

        totals_row_num = None
        if has_col_totals:
            totals_row_num = last_data_row + 1
            # Matches the legacy exporter's convention: the "Total" label sits under
            # the last row-header column, the others in that row stay blank.
            if row_header_depth > 0:
                ws.cell(row=totals_row_num, column=row_header_depth, value="Total")
            for j, cellobj in enumerate(column_totals):
                apply_cell(totals_row_num, row_header_depth + 1 + j, cellobj)
            if has_row_totals and grand_total is not None:
                apply_cell(totals_row_num, row_header_depth + 1 + col_count, grand_total)

        last_row = totals_row_num if totals_row_num else last_data_row
        all_headers = row_headers_text + col_headers_text + (["Total"] if has_row_totals else [])
        tbl = finalize_table(1, last_col, last_row, pTableName or "ExportTable", all_headers)

        if totals_row_num:
            tbl.totalsRowShown = True
            tbl.totalsRowCount = 1
            if row_header_depth > 0:
                tbl.tableColumns[row_header_depth - 1].totalsRowLabel = "Total"
            for j in range(col_count):
                tbl.tableColumns[row_header_depth + j].totalsRowFunction = col_totals_row_function
            if has_row_totals:
                tbl.tableColumns[row_header_depth + col_count].totalsRowFunction = row_totals_row_function
            for c in range(1, last_col + 1):
                ws.cell(row=totals_row_num, column=c).font = HEADER_FONT

        ws.freeze_panes = "A2"

for idx, width in col_widths.items():
    ws.column_dimensions[get_column_letter(idx)].width = min(max(width + 2, 8), 60)

return notes
}

/// Sanitizes+dedupes a candidate Excel sheet name: strips the characters Excel forbids
/// in a sheet name (<code>[ ] : * ? / \</code>), truncates to Excel's 31-character
/// limit, and - given the list of names already used earlier in the same workbook
/// (<var>pUsed</var>, mutated in place is NOT relied on; caller re-appends the result
/// itself) - appends " (2)", " (3)", etc. until the name is unique, re-truncating as
/// needed so the suffixed form still fits in 31 characters.
ClassMethod %SanitizeSheetName(pName, pUsed) [ Language = python, Private ]
{
INVALID_CHARS = set('[]:*?/\\')
base = ''.join(ch for ch in (pName or 'Sheet') if ch not in INVALID_CHARS).strip()
if not base:
    base = 'Sheet'
base = base[:31]

used_lower = set(u.lower() for u in (pUsed or []))
candidate = base
n = 2
while candidate.lower() in used_lower:
    suffix = ' (%d)' % n
    candidate = base[:31 - len(suffix)] + suffix
    n += 1
return candidate
}

}
