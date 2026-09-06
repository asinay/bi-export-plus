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

/// Sets the download filename header, AND - new - resets the response to a plain HTML
/// 503 when <method>BIExport.XLSXWriter.IsAvailable</method> reports Python/openpyxl
/// aren't available. Both have to happen here, not in OnPage() - by the time OnPage()
/// runs, %CSP.Page has already flushed the HTTP header block (built from the CONTENTTYPE
/// class <parameter>Parameter</parameter> above), so neither a %response.SetHeader() nor
/// a %response.ContentType/%response.Status assignment made in OnPage() takes effect
/// (confirmed live both ways: the Content-Disposition header was silently absent despite
/// OnPage() "succeeding" - originally masked because Chrome can't render a spreadsheet
/// mimetype inline and downloads it regardless of the missing header - and, separately,
/// pointing IsAvailable() at a bogus module and setting %response.ContentType/Status from
/// inside OnPage() still produced a 200 with the binary .xlsx content-type on the wire,
/// carrying the HTML error page's bytes - exactly the corrupt-download symptom this
/// feature exists to fix, just re-created by putting the fix in the wrong method).
/// Mirrors how the stock %DeepSee.UI.MDXExcel page sets its own Content-Disposition
/// header in this exact same method, for the same reason.
///
/// NOTE (pre-existing, not fixed here - out of scope for this change): OnPage()'s existing
/// %response.Status assignments for a missing MDX (400) and an unhandled exception (500)
/// have this identical bug and don't reach the client either (confirmed live the same way:
/// a request with no MDX param still comes back 200 with the .xlsx content-type and an
/// empty body) - the fix would be the same shape (decide the status in OnPreHTTP, before
/// %request's other params are known to be well-formed enough to classify), but that's a
/// separate, broader change than this feature's Python-availability check.
ClassMethod OnPreHTTP() As %Boolean
{
    If '##class(BIExport.XLSXWriter).IsAvailable(.tReason) {
        Set %response.ContentType = "text/html"
        Set %response.Status = "503 Service Unavailable"
        Quit 1
    }

    Set tFileName = $Get(%request.Data("FILENAME",1), "export.xlsx")
    Do %response.SetHeader("Content-Disposition", "attachment; filename=""" _ tFileName _ """")
    Quit 1
}

/// Writes a small, legible HTML page reporting that <method>BIExport.XLSXWriter.IsAvailable</method>
/// returned false, with <var>pReason</var> (its own diagnosis of which of the two tiers -
/// Embedded Python not configured, or configured but openpyxl not installed - actually failed)
/// front and center - not <code>ex.DisplayString()</code>, which is a raw exception dump.
/// <method>OnPreHTTP</method> has already reset the response to 503/text-html and skipped
/// setting the Content-Disposition attachment header (so a browser renders this page inline
/// instead of downloading it as a mystery file) by the time this runs.
ClassMethod %WriteUnavailablePage(pReason As %String) [ Private ]
{
    &html<<!DOCTYPE html>
<html>
<head><title>Excel Workbook export unavailable</title>
<style>
  body { font-family: Calibri, "Segoe UI", Arial, sans-serif; margin: 40px; color: #1B2A4A; }
  h1 { font-size: 18px; }
  p.reason { background: #FDEDED; border: 1px solid #F5C2C2; border-radius: 4px; padding: 12px 16px; }
</style>
</head>
<body>
<h1>Excel Workbook export is currently unavailable</h1>
<p class="reason">#($ZCVT(pReason,"O","HTML"))#</p>
</body>
</html>>
}

/// Query parameters (same names/meaning as the stock MDXExcel page where they overlap):
///   MDX or NONCE   - the query text, or a nonce created via CreateQueryNonce() for long MDX
///   TITLE, SUBTITLE, LISTINGNAME, ROWTOTALS (1/0), COLUMNTOTALS (1/0),
///   ROWTOTALAGG, COLUMNTOTALAGG ("sum"/"count"/"min"/"max"/"avg"/"pct"), FILENAME
ClassMethod OnPage() As %Status
{
    Set tSC = $$$OK
    Try {
        // OnPreHTTP() already made the actual response-header decision (too late to do so
        // here - see that method's doc comment); this just re-checks (cheap, no caching -
        // see BIExport.XLSXWriter.IsAvailable()'s own doc comment) to decide whether to
        // write the error body or the real export.
        If '##class(BIExport.XLSXWriter).IsAvailable(.tReason) {
            Do ..%WriteUnavailablePage(tReason)
            Quit
        }

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
