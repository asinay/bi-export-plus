/// Plain CSP page (not a separate REST application) that serves the whole-dashboard
/// multi-sheet .xlsx export, deliberately mirroring <class>BIExport.UI.XLSXPage</class>'s
/// own pattern (itself mirroring the stock <class>%DeepSee.UI.MDXExcel</class> page):
/// living in the SAME CSP application as <class>BIExport.UI.DashboardViewer</class>, the
/// page that opens it via <code>window.open()</code>, so the browser's existing login
/// session is reused automatically instead of hitting the separate-session problem a
/// brand-new <class>%CSP.REST</class> application would (see CLAUDE.md gotcha #3).
Class BIExport.UI.DashboardXLSXPage Extends %CSP.Page
{

Parameter CONTENTTYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/// Sets the download filename header, AND resets the response to a plain HTML 503 when
/// <method>BIExport.XLSXWriter.IsAvailable</method> reports Python/openpyxl aren't
/// available. See <class>BIExport.UI.XLSXPage</class>'s own identical override for the
/// full explanation - by the time OnPage() runs, %CSP.Page has already flushed the HTTP
/// header block built from the CONTENTTYPE parameter above, so neither a
/// %response.SetHeader() nor a %response.ContentType/%response.Status assignment made
/// there takes effect (confirmed live the same way as that class).
ClassMethod OnPreHTTP() As %Boolean
{
    If '##class(BIExport.XLSXWriter).IsAvailable(.tReason) {
        Set %response.ContentType = "text/html"
        Set %response.Status = "503 Service Unavailable"
        Quit 1
    }

    Set tDashboard = $Get(%request.Data("DASHBOARD",1))
    Set tFileName = ..%SuggestFileName(tDashboard)
    Do %response.SetHeader("Content-Disposition", "attachment; filename=""" _ tFileName _ """")
    Quit 1
}

/// Turns a dashboard's full library path ("Folder/Name.dashboard") into a plain
/// download filename ("Name.xlsx") - strips the folder path (portal organization, not
/// part of the report's identity - same reasoning as the Analyzer/widget buttons
/// stripping a pivot's folder prefix from their own filenames) and the ".dashboard"
/// extension, then appends ".xlsx". Falls back to a generic name if DASHBOARD is blank
/// (shouldn't normally happen - OnPage() below 400s on that case - but OnPreHTTP() runs
/// first and has no equivalent bail-out).
ClassMethod %SuggestFileName(pDashboardFullName As %String) As %String [ Private ]
{
    If (pDashboardFullName = "") Quit "dashboard-export.xlsx"
    Set tName = pDashboardFullName
    If (tName [ "/") { Set tName = $Piece(tName,"/",$Length(tName,"/")) }
    If ($ZConvert($Piece(tName,".",$Length(tName,".")),"L") = "dashboard") {
        Set tName = $Piece(tName,".",1,$Length(tName,".")-1)
    }
    Quit tName _ ".xlsx"
}

/// Writes a small, legible HTML page reporting that <method>BIExport.XLSXWriter.IsAvailable</method>
/// returned false. See <class>BIExport.UI.XLSXPage</class>'s own identical method for the full
/// explanation (not <code>ex.DisplayString()</code>, harmless alongside the attachment header
/// OnPreHTTP already set).
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

/// Query parameters: DASHBOARD - the dashboard's full library path (its
/// <property>fullName</property>, e.g. "BIExport Test Dashboard.dashboard" - the same
/// value <class>BIExport.UI.DashboardViewer</class>'s own <property>dashboardName</property>
/// holds).
ClassMethod OnPage() As %Status
{
    Set tSC = $$$OK
    Try {
        // OnPreHTTP() already made the actual response-header decision (too late here -
        // see that method's doc comment); this just re-checks to decide whether to write
        // the error body or the real export.
        If '##class(BIExport.XLSXWriter).IsAvailable(.tReason) {
            Do ..%WriteUnavailablePage(tReason)
            Quit
        }

        Set tDashboard = $Get(%request.Data("DASHBOARD",1))
        If (tDashboard = "") {
            Set %response.Status = "400 Bad Request"
            Write "Missing DASHBOARD parameter"
            Quit
        }

        Set tData = ##class(BIExport.DashboardExport).BuildDashboardExportData(tDashboard, .tExportSC)
        If $$$ISERR(tExportSC) Throw ##class(%Exception.StatusException).CreateFromStatus(tExportSC)

        Set tBytes = ##class(BIExport.XLSXWriter).BuildMultiSheetXLSX(tData.%ToJSON())

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
