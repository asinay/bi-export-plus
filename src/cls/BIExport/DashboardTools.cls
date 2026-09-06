/// Rollout helper for opting existing dashboards into the real .xlsx/JSON export
/// widget (<class>BIExport.UI.Widget.pivot</class>) without hand-editing dashboard XML -
/// same idea as <class>BIExport.Favorites</class> (a helper method instead of a manual
/// edit), but this one rewrites a dashboard's own widget definitions rather than adding
/// a Portal bookmark.
///
/// A widget's rendering class is just a string property on its definition -
/// <property>%DeepSee.Dashboard.Widget::type</property> - confirmed live (dumping
/// <method>%DeepSee.Dashboard.Widget::%CreateVisualWidget</method>'s compiled source)
/// to be instantiated directly by that method, prefixed with
/// "%DeepSee.Component.Widget." unless it already contains a dot. So converting a
/// widget is just: open the dashboard, find the widget by its stable <property>name</property>
/// (not <property>key</property>, which is an unstable hash regenerated on layout
/// changes), set <property>type</property> to "BIExport.UI.Widget.pivot", <method>%Save</method>
/// the dashboard.
///
/// Verified end-to-end live (Playwright) against a scratch clone of
/// "Basic Dashboard Demo.dashboard" before this class was written: cloning via
/// <method>%ConstructClone</method>(1) + re-<method>%Save</method>-ing under a new
/// <property>fullName</property> creates a genuinely independent copy (new %Id), and
/// flipping a widget's <property>type</property> (even to a nonexistent class, as a
/// stress test) + <method>%Save</method> + reload reliably round-trips - the dashboard
/// viewer surfaced a clean in-page error for the bogus class and rendered normally again
/// once the type was fixed, with no corruption either way.
Class BIExport.DashboardTools Extends %RegisteredObject
{

/// Converts one widget (identified by its <property>name</property>, e.g. "Widget1" -
/// see the Widgets tab of the Dashboard Editor, or dump <property>widgets</property> in
/// a Terminal, for the current names) on the dashboard named by
/// <var>pDashboardFullName</var> (its full library path, e.g. "Basic Dashboard Demo.dashboard")
/// to <var>pNewType</var>. Refuses to convert a widget whose current <property>type</property>
/// isn't recognizably in the stock pivot-widget family (bare "pivot", blank, or already
/// <var>pNewType</var>/"BIExport.UI.Widget.pivot") - chart-only/KPI/text widget types
/// have no confirmed compatibility with this class's <method>%OnCreateControls</method>
/// override and aren't a supported target.
ClassMethod ConvertWidget(pDashboardFullName As %String, pWidgetName As %String, pNewType As %String = "BIExport.UI.Widget.pivot") As %Status
{
    Set tSC = $$$OK
    Try {
        Set tDash = ##class(%DeepSee.Dashboard.Utils).%OpenDashboard(pDashboardFullName, .tSC)
        If $$$ISERR(tSC) Quit
        If '$IsObject(tDash) {
            Set tSC = $$$ERROR($$$GeneralError, "Dashboard not found: "_pDashboardFullName)
            Quit
        }

        Set tWidget = ""
        For i=1:1:tDash.widgets.Count() {
            Set w = tDash.widgets.GetAt(i)
            If (w.name = pWidgetName) {
                Set tWidget = w
                Quit
            }
        }
        If '$IsObject(tWidget) {
            Set tSC = $$$ERROR($$$GeneralError, "Widget not found on '"_pDashboardFullName_"': "_pWidgetName)
            Quit
        }

        If '..%IsPivotFamily(tWidget.type, pNewType) {
            Set tSC = $$$ERROR($$$GeneralError, "Widget '"_pWidgetName_"' has type '"_tWidget.type_"', not a recognized pivot-widget type - refusing to convert")
            Quit
        }

        Set tWidget.type = pNewType
        Set tSC = tDash.%Save()
        If $$$ISERR(tSC) Quit

        Write "Converted '"_pWidgetName_"' on '"_pDashboardFullName_"' to "_pNewType, !
    }
    Catch ex {
        Set tSC = ex.AsStatus()
    }
    Quit tSC
}

/// Same as <method>ConvertWidget</method>, but converts every widget on the dashboard
/// whose current <property>type</property> is in the recognized pivot-widget family
/// (see <method>%IsPivotFamily</method>) in one call. Widgets outside that family
/// (chart-only, KPI, text, etc.) are silently skipped rather than erroring the whole
/// call - a dashboard mixing pivot and non-pivot widgets is the normal case, not a
/// mistake.
ClassMethod ConvertAllWidgets(pDashboardFullName As %String, pNewType As %String = "BIExport.UI.Widget.pivot") As %Status
{
    Set tSC = $$$OK
    Try {
        Set tDash = ##class(%DeepSee.Dashboard.Utils).%OpenDashboard(pDashboardFullName, .tSC)
        If $$$ISERR(tSC) Quit
        If '$IsObject(tDash) {
            Set tSC = $$$ERROR($$$GeneralError, "Dashboard not found: "_pDashboardFullName)
            Quit
        }

        Set tConverted = 0
        For i=1:1:tDash.widgets.Count() {
            Set w = tDash.widgets.GetAt(i)
            If ..%IsPivotFamily(w.type, pNewType) {
                Set w.type = pNewType
                Set tConverted = tConverted + 1
            }
        }

        If (tConverted = 0) {
            Write "No pivot-family widgets found on '"_pDashboardFullName_"' - nothing converted", !
            Quit
        }

        Set tSC = tDash.%Save()
        If $$$ISERR(tSC) Quit

        Write "Converted "_tConverted_" widget(s) on '"_pDashboardFullName_"' to "_pNewType, !
    }
    Catch ex {
        Set tSC = ex.AsStatus()
    }
    Quit tSC
}

/// Sets a single widget's <property>type</property> back to the stock "pivot" -
/// the inverse of <method>ConvertWidget</method>, for backing out the change on a
/// dashboard that turns out not to want it.
ClassMethod RevertWidget(pDashboardFullName As %String, pWidgetName As %String) As %Status
{
    Set tSC = $$$OK
    Try {
        Set tDash = ##class(%DeepSee.Dashboard.Utils).%OpenDashboard(pDashboardFullName, .tSC)
        If $$$ISERR(tSC) Quit
        If '$IsObject(tDash) {
            Set tSC = $$$ERROR($$$GeneralError, "Dashboard not found: "_pDashboardFullName)
            Quit
        }

        Set tWidget = ""
        For i=1:1:tDash.widgets.Count() {
            Set w = tDash.widgets.GetAt(i)
            If (w.name = pWidgetName) {
                Set tWidget = w
                Quit
            }
        }
        If '$IsObject(tWidget) {
            Set tSC = $$$ERROR($$$GeneralError, "Widget not found on '"_pDashboardFullName_"': "_pWidgetName)
            Quit
        }

        Set tWidget.type = "pivot"
        Set tSC = tDash.%Save()
        If $$$ISERR(tSC) Quit

        Write "Reverted '"_pWidgetName_"' on '"_pDashboardFullName_"' to stock 'pivot'", !
    }
    Catch ex {
        Set tSC = ex.AsStatus()
    }
    Quit tSC
}

/// True if <var>pType</var> (a widget's current <property>type</property> value) is
/// something this class is willing to convert: blank or bare "pivot" (the stock
/// default, prefixed to %DeepSee.Component.Widget.pivot by %CreateVisualWidget), or
/// already <var>pTargetType</var>/"BIExport.UI.Widget.pivot" (idempotent re-conversion).
/// Anything else - a chart-only, KPI, text widget, or some other custom subclass - is
/// not a confirmed-safe target and is treated as not-pivot-family.
ClassMethod %IsPivotFamily(pType As %String, pTargetType As %String) As %Boolean [ Private ]
{
    Quit (pType = "") || (pType = "pivot") || (pType = pTargetType) || (pType = "BIExport.UI.Widget.pivot")
}

}
