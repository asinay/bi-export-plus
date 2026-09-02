/// Plain CSP page (not a separate REST application) that serves the real .xlsx export,
/// deliberately mirroring how the stock <class>%DeepSee.UI.MDXExcel</class> page works:
/// living in the SAME CSP application as the Analyzer page that links to it, so the
/// browser's existing login session is reused automatically. A separate %CSP.REST
/// application (see <class>BIExport.REST</class>) has its own session/cookie scope and
/// would otherwise force a second login when called from inside an already-open
/// Analyzer tab - this class exists specifically to avoid that.
Include %DeepSee

Class BIExport.UI.XLSXPage Extends %CSP.Page
{

Parameter CONTENTTYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/// Sets the download filename header. Must happen here, not in OnPage() - by the time
/// OnPage() runs, %CSP.Page has already flushed the HTTP header block (built from the
/// CONTENTTYPE parameter above), so a %response.SetHeader() call in OnPage() is too
/// late to take effect (confirmed live: the header was silently absent from the actual
/// response despite the OnPage() call "succeeding" - masked here because Chrome can't
/// render a spreadsheet mimetype inline and downloads it regardless, unlike JSON).
/// Mirrors how the stock %DeepSee.UI.MDXExcel page sets this exact header in this exact
/// same method, for the same reason.
ClassMethod OnPreHTTP() As %Boolean
{
    Set tFileName = $Get(%request.Data("FILENAME",1), "export.xlsx")
    Do %response.SetHeader("Content-Disposition", "attachment; filename=""" _ tFileName _ """")
    Quit 1
}

/// Query parameters (same names/meaning as the stock MDXExcel page where they overlap):
///   MDX or NONCE   - the query text, or a nonce created via CreateQueryNonce() for long MDX
///   TITLE, SUBTITLE, LISTINGNAME, ROWTOTALS (1/0), COLUMNTOTALS (1/0),
///   ROWTOTALAGG, COLUMNTOTALAGG ("sum"/"count"/"min"/"max"/"avg"/"pct"), FILENAME
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
        Set tOptions.rowTotals = ($Get(%request.Data("ROWTOTALS",1)) = "1")
        Set tOptions.columnTotals = ($Get(%request.Data("COLUMNTOTALS",1)) = "1")
        Set tOptions.rowTotalAgg = $Get(%request.Data("ROWTOTALAGG",1))
        Set tOptions.columnTotalAgg = $Get(%request.Data("COLUMNTOTALAGG",1))
        Set tOptions.rowCaptions = $Get(%request.Data("ROWCAPTIONS",1))
        Set tOptions.filters = ##class(BIExport.MDXExport).%ParseFilterPairs($Get(%request.Data("FILTERNAMES",1)), $Get(%request.Data("FILTERVALUES",1)))

        Set tData = ##class(BIExport.MDXExport).BuildExportData(tMDX, tOptions)
        Set tBytes = ##class(BIExport.XLSXWriter).BuildXLSX(tData.%ToJSON())

        Set tStream = ##class(%Stream.TmpBinary).%New()
        Do tStream.Write(tBytes)
        Do tStream.Rewind()
        Do tStream.OutputToDevice()
    }
    Catch ex {
        Set %response.Status = "500 Internal Server Error"
        Write ex.DisplayString()
    }
    Quit $$$OK
}

}
