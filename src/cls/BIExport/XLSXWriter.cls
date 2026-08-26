/// Renders the normalized JSON structure produced by <class>BIExport.MDXExport</class>
/// into a genuine .xlsx workbook (real cell types and number formats - not an HTML
/// table saved with an .xls extension).
///
/// Layout: the data sheet is a real Excel structured Table (native autofilter/banded
/// rows, sortable, recognized as a table by Excel/other tools) starting at A1 - the
/// header row is never pushed down by metadata. Title/subtitle/filters instead go on
/// a separate "Info" sheet, plus the workbook's own document properties.
Class BIExport.XLSXWriter Extends %RegisteredObject
{

/// pJSON is the %ToJSON() output of a BIExport.MDXExport.BuildExportData() result.
/// Returns the raw bytes of an .xlsx file.
ClassMethod BuildXLSX(pJSON As %String) As %Binary [ Language = python ]
{
import json
import io
import datetime
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

data = json.loads(pJSON)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Export"

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

# ---------------------------------------------------------------------------
# Info sheet + document properties - title/subtitle/filters live here, never
# inline on the data sheet, so the data sheet's header always starts at A1.
# ---------------------------------------------------------------------------
title = data.get("title") or ""
subtitle = data.get("subtitle") or ""
cube = data.get("cube") or ""
filters = data.get("filters") or []

info = wb.create_sheet("Info")
r = 1
if title:
    c = info.cell(row=r, column=1, value="Report")
    c.font = HEADER_FONT
    info.cell(row=r, column=2, value=title)
    r += 1
if cube:
    c = info.cell(row=r, column=1, value="Cube")
    c.font = HEADER_FONT
    info.cell(row=r, column=2, value=cube)
    r += 1
if subtitle:
    c = info.cell(row=r, column=1, value="Subtitle")
    c.font = HEADER_FONT
    info.cell(row=r, column=2, value=subtitle)
    r += 1
for f in filters:
    c = info.cell(row=r, column=1, value=f.get("name", ""))
    c.font = HEADER_FONT
    info.cell(row=r, column=2, value=f.get("value", ""))
    r += 1
if r == 1:
    info.cell(row=r, column=1, value="(no title, cube, subtitle, or filters supplied for this export)")
info_next_row = r
info.column_dimensions["A"].width = 24
info.column_dimensions["B"].width = 50

if title:
    wb.properties.title = title
if subtitle:
    wb.properties.description = subtitle
if cube:
    wb.properties.subject = cube

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
        finalize_table(1, len(headers), last_row, "ExportTable", headers)
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

        info.cell(row=info_next_row, column=1,
                   value="Note: this pivot has a dimension on Columns (multi-level column "
                         "header), so this sheet keeps a true merged header instead of a "
                         "flattened Excel Table - no autofilter/banded rows on this export.")
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
        tbl = finalize_table(1, last_col, last_row, "ExportTable", all_headers)

        if totals_row_num:
            tbl.totalsRowShown = True
            tbl.totalsRowCount = 1
            if row_header_depth > 0:
                tbl.tableColumns[row_header_depth - 1].totalsRowLabel = "Total"
            for j in range(col_count):
                tbl.tableColumns[row_header_depth + j].totalsRowFunction = "sum"
            if has_row_totals:
                tbl.tableColumns[row_header_depth + col_count].totalsRowFunction = "sum"
            for c in range(1, last_col + 1):
                ws.cell(row=totals_row_num, column=c).font = HEADER_FONT

        ws.freeze_panes = "A2"

for idx, width in col_widths.items():
    ws.column_dimensions[get_column_letter(idx)].width = min(max(width + 2, 8), 60)

buf = io.BytesIO()
wb.save(buf)
return buf.getvalue()
}

}
