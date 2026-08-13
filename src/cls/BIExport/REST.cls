/// REST service exposing real-file BI exports (starting with genuine .xlsx,
/// replacing the Analyzer's HTML-saved-as-.xls behavior). See CLAUDE.md for
/// the design rationale and docs/iris-environment.md for how this is deployed.
Class BIExport.REST Extends %CSP.REST
{

Parameter HandleCorsRequest = 0;

XData UrlMap
{
<Routes>
<Route Url="/xlsx" Method="POST" Call="ExportXLSX"/>
<Route Url="/json" Method="POST" Call="ExportJSON"/>
</Routes>
}

/// POST body (JSON): {
///   "mdx": "<MDX query text>",         - required
///   "filename": "report.xlsx",         - optional, defaults to "<title>.xlsx", or
///                                        "<cube>_<timestamp>.xlsx" if no title
///   "title": "...", "subtitle": "...", - optional, shown on the Info sheet
///   "cube": "HoleFoods Sales",         - optional, shown on the Info sheet even with
///                                        no title, and used in the default filename
///   "listing": "ListingName",          - optional, for drillthrough/listing MDX
///   "totals": true,                    - optional, include row/column totals
///   "filters": [{"name":"...","value":"..."}]  - optional, recap block on the Info sheet
///   "rowCaptions": ["YearSold", ...]   - optional, row axis dimension captions (outermost
///                                        first) for the row-header columns; the MDX/
///                                        ResultSet API can't derive these on its own -
///                                        this is metadata an external caller would need
///                                        to already know from how it built the MDX
/// }
ClassMethod ExportXLSX() As %Status
{
    Set tSC = $$$OK
    Try {
        Set tBody = %request.Content
        Do tBody.Rewind()
        Set tJSONIn = tBody.Read(tBody.Size)

        Set tReq = {}.%FromJSON(tJSONIn)
        Set tMDX = tReq.%Get("mdx", "")
        If (tMDX = "") {
            Set %response.Status = "400 Bad Request"
            Write "Missing required 'mdx' field"
            Quit
        }

        Set tData = ##class(BIExport.MDXExport).BuildExportData(tMDX, tReq)
        Set tJSON = tData.%ToJSON()
        Set tBytes = ##class(BIExport.XLSXWriter).BuildXLSX(tJSON)

        Set tFileName = tReq.%Get("filename", ..%DefaultFileName(tReq, "xlsx"))

        Set %response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        Do %response.SetHeader("Content-Disposition", "attachment; filename=""" _ tFileName _ """")

        Set tStream = ##class(%Stream.TmpBinary).%New()
        Do tStream.Write(tBytes)
        Do tStream.Rewind()
        Do tStream.OutputToDevice()
    }
    Catch ex {
        Set %response.Status = "500 Internal Server Error"
        Set %response.ContentType = "text/plain"
        Write ex.DisplayString()
    }
    Quit $$$OK
}

/// POST body (JSON): same fields as /xlsx (mdx, filename, title, subtitle, listing,
/// totals, filters) - returns BIExport.MDXExport's normalized structure as-is. No
/// Python involved: %DynamicObject already has %ToJSON(), so this is pure ObjectScript.
ClassMethod ExportJSON() As %Status
{
    Set tSC = $$$OK
    Try {
        Set tBody = %request.Content
        Do tBody.Rewind()
        Set tJSONIn = tBody.Read(tBody.Size)

        Set tReq = {}.%FromJSON(tJSONIn)
        Set tMDX = tReq.%Get("mdx", "")
        If (tMDX = "") {
            Set %response.Status = "400 Bad Request"
            Write "Missing required 'mdx' field"
            Quit
        }

        Set tData = ##class(BIExport.MDXExport).BuildExportData(tMDX, tReq)

        Set tFileName = tReq.%Get("filename", ..%DefaultFileName(tReq, "json"))

        Set %response.ContentType = "application/json"
        Do %response.SetHeader("Content-Disposition", "attachment; filename=""" _ tFileName _ """")

        Do tData.%ToJSON()
    }
    Catch ex {
        Set %response.Status = "500 Internal Server Error"
        Set %response.ContentType = "text/plain"
        Write ex.DisplayString()
    }
    Quit $$$OK
}

/// Filename to use when the caller didn't supply one: "<title>.<ext>" if a title
/// was given, else "<cube or 'export'>_<short timestamp>.<ext>" - matches
/// BIExport.UI.Analyzer.buildExportURL()'s fallback for an ad-hoc pivot with no
/// saved name to hang a filename on, so repeated title-less exports don't all
/// collide on a single literal "export.xlsx".
ClassMethod %DefaultFileName(pReq As %DynamicObject, pExt As %String) As %String [ Private ]
{
    Set tTitle = pReq.%Get("title", "")
    If (tTitle '= "") {
        Quit tTitle _ "." _ pExt
    }
    Set tCube = pReq.%Get("cube", "")
    Set tBase = $Select(tCube'="":tCube,1:"export")
    Set tTS = $TR($ZDate($H,8),"-") _ "-" _ $Extract($TR($ZTime($Piece($H,",",2),1),":"),1,4)
    Quit tBase _ "_" _ tTS _ "." _ pExt
}

}
