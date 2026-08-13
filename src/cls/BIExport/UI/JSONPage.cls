/// Plain CSP page (not a separate REST application) that serves the real JSON export,
/// mirroring <class>BIExport.UI.XLSXPage</class>'s reasoning exactly: living in the SAME
/// CSP application as the Analyzer page that links to it, so the browser's existing
/// login session is reused automatically instead of hitting <class>BIExport.REST</class>'s
/// separate session/cookie scope.
Include %DeepSee

Class BIExport.UI.JSONPage Extends %CSP.Page
{

Parameter CONTENTTYPE = "application/json";

/// Sets the download filename header. Must happen here, not in OnPage() - by the time
/// OnPage() runs, %CSP.Page has already flushed the HTTP header block (built from the
/// CONTENTTYPE parameter above), so a %response.SetHeader() call in OnPage() is too
/// late to take effect. Confirmed by checking how the stock %DeepSee.UI.MDXExcel page
/// does exactly this same thing, in this exact same method, for the same reason.
ClassMethod OnPreHTTP() As %Boolean
{
    Set tFileName = $Get(%request.Data("FILENAME",1), "export.json")
    Do %response.SetHeader("Content-Disposition", "attachment; filename=""" _ tFileName _ """")
    Quit 1
}

/// Query parameters (same names/meaning as <class>BIExport.UI.XLSXPage</class>):
///   MDX or NONCE   - the query text, or a nonce created via CreateQueryNonce() for long MDX
///   TITLE, SUBTITLE, LISTINGNAME, TOTALS (1/0), FILENAME
ClassMethod OnPage() As %Status
{
    Set tSC = $$$OK
    Try {
        Set tNonce = $Get(%request.Data("NONCE",1))
        If (tNonce '= "") {
            Set tMDX = $Get($$$DeepSeeQueryNonceGLVN(tNonce))
            Kill $$$DeepSeeQueryNonceGLVN(tNonce)
        } Else {
            Set tMDX = $Get(%request.Data("MDX",1))
        }
        If (tMDX = "") {
            Set %response.Status = "400 Bad Request"
            Quit
        }

        Set tOptions = {}
        Set tOptions.title = $Get(%request.Data("TITLE",1))
        Set tOptions.subtitle = $Get(%request.Data("SUBTITLE",1))
        Set tOptions.cube = $Get(%request.Data("CUBE",1))
        Set tOptions.listing = $Get(%request.Data("LISTINGNAME",1))
        Set tOptions.totals = ($Get(%request.Data("TOTALS",1)) = "1")
        Set tOptions.rowCaptions = $Get(%request.Data("ROWCAPTIONS",1))
        Set tOptions.filters = ##class(BIExport.MDXExport).%ParseFilterPairs($Get(%request.Data("FILTERNAMES",1)), $Get(%request.Data("FILTERVALUES",1)))

        Set tData = ##class(BIExport.MDXExport).BuildExportData(tMDX, tOptions)

        Do tData.%ToJSON()
    }
    Catch ex {
        Set %response.Status = "500 Internal Server Error"
        Write ex.DisplayString()
    }
    Quit $$$OK
}

}
