/// Adds a "Favorite" bookmark in the IRIS BI User Portal (the same Favorites list
/// shown on the Analyzer/Home page) that opens a saved pivot through
/// <class>BIExport.UI.Analyzer</class> instead of the stock <class>%DeepSee.UI.Analyzer</class> -
/// so the Export dropdown's "Excel Workbook"/"JSON" options are there without needing
/// to change any existing link/shortcut URL. Implemented as a <class>%DeepSee.UserLibrary.Link</class>
/// folder item (the same mechanism the Portal's own "New Link" dialog uses), favorited
/// for the current user via <class>%DeepSee.UserLibrary.Utils</class>.
///
/// See README.md, "Add a Favorites bookmark instead" for usage from the Analyzer's
/// own "Show MDX" pivot name, or call directly from a Terminal:
/// <example>
/// Do ##class(BIExport.Favorites).AddAnalyzerFavorite("HoleFoods/By Region.pivot")
/// </example>
Class BIExport.Favorites Extends %RegisteredObject
{

/// Creates (or updates) a <class>%DeepSee.UserLibrary.Link</class> pointing at
/// <class>BIExport.UI.Analyzer</class> for the saved pivot named by <var>pPivotFullName</var>
/// (its full library path, e.g. "HoleFoods/By Region.pivot" - the same value shown as
/// "fullName" in the Portal's pivot list), and adds it to the current user's ($UserName)
/// Favorites. <var>pFolderName</var> controls which folder the bookmark itself is filed
/// under (created if missing); <var>pTitle</var> defaults to the pivot's own item name.
ClassMethod AddAnalyzerFavorite(pPivotFullName As %String, pFolderName As %String = "BIExport", pTitle As %String = "") As %Status
{
    Set tSC = $$$OK
    Try {
        If (pPivotFullName = "") {
            Set tSC = $$$ERROR($$$GeneralError, "pPivotFullName is required, e.g. ""HoleFoods/By Region.pivot""")
            Quit
        }
        Set tSC = ##class(%DeepSee.UserLibrary.FolderItem).%ParseFullName(pPivotFullName, .tFolder, .tItemName, .tExt)
        If $$$ISERR(tSC) Quit
        Set:pTitle="" pTitle = tItemName

        Set tHref = "BIExport.UI.Analyzer.cls?PIVOT=" _ $ZCVT($ZCVT(pPivotFullName, "O", "UTF8"), "O", "URL")

        Set tLink = ##class(%DeepSee.UserLibrary.Link).%New()
        Set tLink.fullName = pFolderName _ "/" _ pTitle
        Set tLink.href = tHref
        Set tLink.title = pTitle
        Set tLink.description = "Opens '" _ pPivotFullName _ "' via BIExport.UI.Analyzer.cls (adds Excel Workbook/JSON export options to the Export dropdown)."
        Set tLink.owner = $UserName
        Set tLink.public = 1
        Set tSC = tLink.%Save()
        If $$$ISERR(tSC) Quit

        Set tSC = ##class(%DeepSee.UserLibrary.Utils).%AddFavorite(tLink.fullName)
        If $$$ISERR(tSC) Quit

        Write "Added favorite: ", tLink.fullName, " -> ", tHref, !
    }
    Catch ex {
        Set tSC = ex.AsStatus()
    }
    Quit tSC
}

/// Removes the bookmark from the current user's Favorites list (does not delete the
/// underlying <class>%DeepSee.UserLibrary.Link</class> item - same as unfavoriting
/// from the Portal UI). Pass the same <var>pFolderName</var>/<var>pTitle</var> used
/// with <method>AddAnalyzerFavorite</method> (or the exact bookmark fullName as
/// <var>pTitle</var> with <var>pFolderName</var>="").
ClassMethod RemoveAnalyzerFavorite(pFolderName As %String, pTitle As %String) As %Status
{
    Set tLink = ##class(%DeepSee.UserLibrary.Link).%New()
    Set tLink.fullName = $S(pFolderName="":"", 1:pFolderName_"/") _ pTitle
    Quit ##class(%DeepSee.UserLibrary.Utils).%RemoveFavorite(tLink.fullName)
}

}
