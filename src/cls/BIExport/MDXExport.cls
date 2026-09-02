/// Executes an MDX query and flattens the resulting cellset (or drillthrough listing)
/// into a normalized, JSON-friendly structure that <class>BIExport.XLSXWriter</class>
/// (or any other renderer) can consume without needing to know anything about
/// %DeepSee.ResultSet. Mirrors the cellset-walking logic used by the stock
/// Analyzer export buttons (%DeepSee.UI.MDXExcel / %DeepSee.Component.deepseeCSVProvider)
/// but keeps real values/types instead of pre-formatting everything as HTML/CSV text.
Class BIExport.MDXExport Extends %RegisteredObject
{

/// pOptions supported keys (all optional):
///   title, subtitle       - %String, printed above the table
///   cube                  - %String, cube name/caption - shown on the Info sheet even
///                            when title is blank (e.g. an ad-hoc pivot that was never
///                            saved, so has no name to show)
///   listing               - %String, listing name to use if the MDX is a drillthrough/listing query
///   rowTotals             - %Boolean, include a "Total" column (one value per row, summed
///                            across columns) - default false. Independent of columnTotals;
///                            a pivot can have either, both, or neither turned on, and this
///                            must mirror that exactly rather than OR-ing them into one flag
///                            (doing so was a bug - see BIExport.UI.Analyzer's buildExportURL()).
///   columnTotals          - %Boolean, include a "Total" row (one value per column, summed
///                            down rows) - default false. Independent of rowTotals.
///   rowTotalAgg           - %String, aggregation function for rowTotals - one of "sum"
///                            (default), "count", "min", "max", "avg", "pct" (% of total).
///                            Passed straight through to %DeepSee.ResultSet's %GetRowTotal -
///                            these are exactly the values its pAgg parameter accepts, and
///                            also exactly the <option value="..."> values in the Pivot
///                            Options dialog's Summary aggregation dropdown, so no translation
///                            is needed at either end.
///   columnTotalAgg        - %String, aggregation function for columnTotals - same values
///                            as rowTotalAgg, passed to %GetColumnTotal.
///   filters               - array of {name, value} objects, printed as a recap block
///   rowCaptions            - the row axis's dimension/level captions (e.g. "YearSold",
///                            "Product Category"), outermost first. The ResultSet API has
///                            no way to derive these itself - they're metadata the pivot
///                            table widget already has client-side (this is exactly what
///                            the stock exportExcel() passes as ROWCAPTIONS, sourced from
///                            table.getRowCaptions().join('\n') - see BIExport.UI.Analyzer's
///                            buildExportURL()). Accepts either a %DynamicArray of strings
///                            or a single "\n"-delimited %String in that same format.
ClassMethod BuildExportData(pMDX As %String, pOptions As %DynamicObject = "") As %DynamicObject
{
    If '$IsObject(pOptions) { Set pOptions = {} }
    Set tResult = {}
    Set tResult."title" = ..%StrOrNull(pOptions.title)
    Set tResult."subtitle" = ..%StrOrNull(pOptions.subtitle)
    Set tResult."cube" = ..%StrOrNull(pOptions.cube)
    If (pOptions.filters = "") {
        Set tResult.filters = []
    } Else {
        Set tResult.filters = pOptions.filters
    }
    Set tRowCaptions = ..%ParseCaptions(pOptions.%Get("rowCaptions"))

    Set tRS = ##class(%DeepSee.ResultSet).%New()
    Set tSC = tRS.%PrepareMDX(pMDX)
    If $$$ISERR(tSC) Throw ##class(%Exception.StatusException).CreateFromStatus(tSC)

    Set tListingName = pOptions.%Get("listing", "")
    If (tListingName '= "") {
        Set tSC = tRS.%SetListing(tListingName)
        If $$$ISERR(tSC) Throw ##class(%Exception.StatusException).CreateFromStatus(tSC)
    }
    Set tRS.%Query.%MaxListingRows = 0

    Set tSC = tRS.%ExecuteAsynch(,,1)
    If $$$ISERR(tSC) Throw ##class(%Exception.StatusException).CreateFromStatus(tSC)

    If tRS.%IsDrillThrough() {
        Set tResult.kind = "listing"
        Do ..%BuildListing(tRS, tResult)
    } Else {
        Set tResult.kind = "cellset"
        Set tRowTotals = $Select($IsObject(pOptions):(pOptions.rowTotals=1),1:0)
        Set tColTotals = $Select($IsObject(pOptions):(pOptions.columnTotals=1),1:0)
        Set tRowTotalAgg = $Select($IsObject(pOptions)&&(pOptions.%Get("rowTotalAgg","")'=""):pOptions.rowTotalAgg,1:"sum")
        Set tColTotalAgg = $Select($IsObject(pOptions)&&(pOptions.%Get("columnTotalAgg","")'=""):pOptions.columnTotalAgg,1:"sum")
        Do ..%BuildCellset(tRS, tResult, tRowTotals, tColTotals, tRowTotalAgg, tColTotalAgg)

        Set tRowHeaderNames = []
        For n=1:1:tResult.rowHeaderDepth {
            Do tRowHeaderNames.%Push($Select(n<=tRowCaptions.Count():tRowCaptions.GetAt(n),1:""))
        }
        Set tResult.rowHeaderNames = tRowHeaderNames
    }

    Quit tResult
}

/// Builds a filters array (in the {name,value} shape BuildExportData's "filters"
/// option expects) from two parallel "\n"-delimited strings - the same shape
/// XLSXPage.cls/JSONPage.cls receive via FILTERNAMES/FILTERVALUES query params,
/// sourced from the pivot table widget's own table.getFilterInfo() client-side
/// (mirroring what stock exportExcel() does with those same two arrays).
ClassMethod %ParseFilterPairs(pNames As %String, pValues As %String) As %DynamicArray
{
    Set tResult = []
    If (pNames = "") Quit tResult
    Set tCount = $Length(pNames,$C(10))
    For i=1:1:tCount {
        Do tResult.%Push({"name": ($Piece(pNames,$C(10),i)), "value": ($Piece(pValues,$C(10),i))})
    }
    Quit tResult
}

/// Normalizes pValue (either a %DynamicArray of strings, or a single "\n"-delimited
/// %String) into a %ListOfDataTypes of caption strings, outermost first.
ClassMethod %ParseCaptions(pValue) As %ListOfDataTypes [ Private ]
{
    Set tList = ##class(%ListOfDataTypes).%New()
    If $IsObject($Get(pValue)) {
        For i=0:1:pValue.%Size()-1 { Do tList.Insert(pValue.%Get(i)) }
    } ElseIf ($Get(pValue) '= "") {
        For i=1:1:$Length(pValue,$C(10)) { Do tList.Insert($Piece(pValue,$C(10),i)) }
    }
    Quit tList
}

ClassMethod %StrOrNull(pValue) As %String
{
    Quit $Select($IsObject($Get(pValue)):"",1:$Get(pValue,""))
}

/// Flattens a drillthrough/listing result into {columns:[{name,numeric}], rows:[[...]]}
ClassMethod %BuildListing(pRS As %DeepSee.ResultSet, pResult As %DynamicObject) [ Private ]
{
    Set tSC = pRS.%GetListingResultSet(.tListingRS, .tListingFields)
    If $$$ISERR(tSC) Throw ##class(%Exception.StatusException).CreateFromStatus(tSC)

    Set tColumnMeta = tListingRS.%GetMetaData().columns
    Set tColumns = []
    If (tListingFields = "") {
        Set tColCount = tColumnMeta.Count()
        For c=1:1:tColCount {
            Set tInfo = tColumnMeta.GetAt(c)
            Set tNumeric = $Case(tInfo.ODBCType, 2:1, 4:1, 8:1, :0)
            Do tColumns.%Push({"name": (tInfo.colName), "numeric": (tNumeric)})
        }
    } Else {
        Set tColCount = $ListLength(tListingFields)
        For c=1:1:tColCount {
            Do tColumns.%Push({"name": ($ListGet(tListingFields,c)), "numeric": 0})
        }
    }
    Set pResult.columns = tColumns

    Set tRows = []
    While tListingRS.%Next(.tSC) {
        If $$$ISERR(tSC) Throw ##class(%Exception.StatusException).CreateFromStatus(tSC)
        Set tRow = []
        For c=1:1:tColCount {
            If (tListingFields = "") {
                Set tValue = tListingRS.%GetData(c)
            } Else {
                Set tValue = tListingRS.%Get($ListGet(tListingFields,c))
            }
            Do tRow.%Push(tValue)
        }
        Do tRows.%Push(tRow)
    }
    Set pResult.rows = tRows
}

/// Flattens an MDX cellset into headers + typed data cells (+ optional totals).
/// pRowTotals and pColTotals are independent - see BuildExportData's doc comment.
ClassMethod %BuildCellset(pRS As %DeepSee.ResultSet, pResult As %DynamicObject, pRowTotals As %Boolean, pColTotals As %Boolean, pRowTotalAgg As %String = "sum", pColTotalAgg As %String = "sum") [ Private ]
{
    Set tAxisCount = pRS.%GetAxisCount()

    // Single-value answer (no axes) - report it as a 1x1 table.
    If (tAxisCount = 0) {
        Set tValue = pRS.%GetValue()
        Set pResult.colHeaderDepth = 1
        Set pResult.rowHeaderDepth = 1
        Set pResult.colCount = 1
        Set pResult.rowCount = 1
        Set pResult.columnHeaders = [["All"]]
        Set pResult.rowHeaders = [["Count"]]
        Set tCell = ..%MakeNumericCell(tValue, "")
        Set tRowArr = []
        Do tRowArr.%Push(tCell)
        Set tData = []
        Do tData.%Push(tRowArr)
        Set pResult.cells = tData
        Quit
    }

    For a=1:1:tAxisCount { Set tAxisSize(a) = pRS.%GetAxisSize(a) }
    Set tColCount = +$Get(tAxisSize(1))
    Set tHasRows = (tAxisCount >= 2)
    Set tRowCount = $Select(tHasRows:+$Get(tAxisSize(2)),1:1)

    If (tColCount = 0) || (tRowCount = 0) {
        Set pResult.colHeaderDepth = 0
        Set pResult.rowHeaderDepth = 0
        Set pResult.colCount = 0
        Set pResult.rowCount = 0
        Set pResult.columnHeaders = []
        Set pResult.rowHeaders = []
        Set pResult.cells = []
        Quit
    }

    // First pass: column headers + per-column format/solve order.
    Set tColHeaderDepth = 1
    For c=1:1:tColCount {
        Kill tNames
        Set tCount = pRS.%GetOrdinalLabel(.tNames,1,c,.tFormat,.tSolveOrder)
        Set:(tCount>tColHeaderDepth) tColHeaderDepth = tCount
        Set tColFormat(c) = tFormat
        Set tColSolveOrder(c) = tSolveOrder
        For n=1:1:tCount { Set tColLabel(c,n) = $Get(tNames(n)) }
    }

    // Second pass: row headers + per-row format/solve order.
    Set tRowHeaderDepth = 1
    If (tHasRows) {
        For r=1:1:tRowCount {
            Kill tLabels
            Set tCount = pRS.%GetOrdinalLabel(.tLabels,2,r,.tFormat,.tSolveOrder)
            Set:(tCount>tRowHeaderDepth) tRowHeaderDepth = tCount
            Set tRowFormat(r) = tFormat
            Set tRowSolveOrder(r) = tSolveOrder
            For n=1:1:tCount { Set tRowLabel(r,n) = $Get(tLabels(n)) }
        }
    } Else {
        Set tRowFormat(1) = ""
        Set tRowSolveOrder(1) = 0
        Set tRowLabel(1,1) = "Results"
    }

    Set tDefaultFormat = pRS.%GetDefaultFormat()
    Set tHasOverride = pRS.%HasCellFormatOverride()

    // Column headers, outermost (depth tColHeaderDepth) first, innermost (1) last - top-to-bottom.
    // Pushed with explicit type "string" - member captions that happen to look numeric
    // (e.g. a Year value "2024", or a flag member "0"/"1") would otherwise get
    // auto-detected as JSON numbers, producing a mixed string/number array that crashes
    // XLSXWriter's header-flattening join for any pivot with a ragged/asymmetric column axis.
    Set tColumnHeaders = []
    For lvl=tColHeaderDepth:-1:1 {
        Set tRowOfHeaders = []
        For c=1:1:tColCount {
            Do tRowOfHeaders.%Push($Get(tColLabel(c,lvl)),"string")
        }
        Do tColumnHeaders.%Push(tRowOfHeaders)
    }
    Set pResult.columnHeaders = tColumnHeaders

    // Row headers, outermost (depth) first, innermost (1) last - left-to-right per row.
    Set tRowHeaders = []
    For r=1:1:tRowCount {
        Set tHeaderRow = []
        For lvl=tRowHeaderDepth:-1:1 {
            Do tHeaderRow.%Push($Get(tRowLabel(r,lvl)),"string")
        }
        Do tRowHeaders.%Push(tHeaderRow)
    }
    Set pResult.rowHeaders = tRowHeaders

    Set pResult.colHeaderDepth = tColHeaderDepth
    Set pResult.rowHeaderDepth = tRowHeaderDepth
    Set pResult.colCount = tColCount
    Set pResult.rowCount = tRowCount

    // Data cells.
    Set tData = []
    For r=1:1:tRowCount {
        Set tRowArr = []
        For c=1:1:tColCount {
            Set tValue = $Case(tAxisCount, 1:pRS.%GetOrdinalValue(c), :pRS.%GetOrdinalValue(c,r))

            If (+$Get(tColSolveOrder(c)) >= +$Get(tRowSolveOrder(r))) {
                Set tFormat = $Select($Get(tColFormat(c))'="":tColFormat(c), $Get(tRowFormat(r))'="":tRowFormat(r), 1:tDefaultFormat)
            } Else {
                Set tFormat = $Select($Get(tRowFormat(r))'="":tRowFormat(r), $Get(tColFormat(c))'="":tColFormat(c), 1:tDefaultFormat)
            }
            If ((tAxisCount>1) && tHasOverride) {
                Set tFO = pRS.%GetCellFormatOverride(c,r)
                Set:tFO'="" tFormat = tFO
            }

            Do tRowArr.%Push(..%MakeCell(tValue, tFormat))
        }
        Do tData.%Push(tRowArr)
    }
    Set pResult.cells = tData

    If ('pRowTotals) && ('pColTotals) Quit

    // Column totals: one value per column, summed down the rows - rendered as a "Total" row.
    // "pct" (% of Total) returns a raw 0-1 fraction, not a value in the column's own
    // units, so it needs its own percent format rather than the column's normal one.
    If (pColTotals) {
        Set tColTotals = []
        Set tColPct = (pColTotalAgg="pct")
        For c=1:1:tColCount {
            Set tFormat = $Select(tColPct:"0.00%",$Get(tColFormat(c))'="":tColFormat(c),1:tDefaultFormat)
            Do tColTotals.%Push(..%MakeCell(pRS.%GetColumnTotal(c,pColTotalAgg),tFormat))
        }
        Set pResult.columnTotals = tColTotals
        Set pResult.columnTotalAgg = pColTotalAgg
    }

    // Row totals: one value per row, summed across columns - rendered as a "Total" column.
    // Grand total is only meaningful (and only used by the writer) alongside row totals.
    If (pRowTotals) {
        Set tRowTotals = []
        Set tGrandTotal = 0
        Set tRowPct = (pRowTotalAgg="pct")
        For r=1:1:tRowCount {
            Set tFormat = $Select(tRowPct:"0.00%",$Get(tRowFormat(r))'="":tRowFormat(r),1:tDefaultFormat)
            Set tSum = pRS.%GetRowTotal(r,pRowTotalAgg)
            Set tGrandTotal = tGrandTotal + tSum
            Do tRowTotals.%Push(..%MakeCell(tSum,tFormat))
        }
        Set pResult.rowTotals = tRowTotals
        Set pResult.rowTotalAgg = pRowTotalAgg
        Set pResult.grandTotal = ..%MakeCell(tGrandTotal,$Select(tRowPct:"0.00%",1:tDefaultFormat))
    }
}

/// Builds a typed cell object {t, v, fmt} from a raw MDX cell value + MDX format string.
ClassMethod %MakeCell(pValue, pFormat As %String) As %DynamicObject [ Private ]
{
    If (pFormat = "%string%") {
        Quit {"t":"s","v":(pValue),"fmt":"@"}
    }
    ElseIf (pFormat = "%date%") {
        If (pValue = "") {
            Quit {"t":"s","v":""}
        }
        Set tISO = $ZDT(pValue\1,3)
        Quit {"t":"d","v":(tISO),"fmt":"mm/dd/yyyy"}
    }
    ElseIf (pFormat = "%time%") {
        Quit {"t":"t","v":(+pValue),"fmt":"[h]:mm:ss"}
    }
    Else {
        Quit ..%MakeNumericCell(pValue, pFormat)
    }
}

ClassMethod %MakeNumericCell(pValue, pFormat As %String) As %DynamicObject [ Private ]
{
    If $IsValidNum(pValue) {
        Quit {"t":"n","v":(+pValue),"fmt":(..%TranslateFormat(pFormat))}
    }
    Quit {"t":"s","v":(pValue),"fmt":"@"}
}

/// Translates an MDX/ODBC-style number format string (e.g. "#,#0.00", "$#,#0") into a
/// standard Excel number_format string. Real xlsx formats don't need the backslash
/// escaping the legacy HTML "mso-number-format" trick required.
ClassMethod %TranslateFormat(pFormat As %String) As %String [ Private ]
{
    If (pFormat = "") Quit "#,##0"

    Set tHasPercent = (pFormat["%")
    Set tHasComma = (pFormat[",")
    Set tHasDollar = (pFormat["$")

    Set tScale = 0
    Set tDot = $Find(pFormat,".")
    If (tDot) {
        Set p = tDot
        While $Case($Extract(pFormat,p),"#":1,"0":1,:0) {
            Set tScale = tScale + 1
            Set p = p + 1
        }
    }

    Set tInt = $Select(tHasComma:"#,##0",1:"0")
    Set tOut = $Select(tScale>0:tInt_"."_$Translate($Justify("",tScale)," ","0"),1:tInt)
    Set:tHasDollar tOut = "$"_tOut
    Set:tHasPercent tOut = tOut_"%"
    Quit tOut
}

}
