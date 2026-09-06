# Feature request: native real-file export for IRIS BI (Analyzer & Dashboards)

Draft content for a JIRA ticket to InterSystems Development/Product Management,
requesting that IRIS BI/DeepSee ship the export capability this module
(`bi-export-plus`) currently has to bolt on as a third-party add-on. Everything
below is written from the customer/requester side, with a recommendation on
which pieces to actually ask for.

## Problem

The stock Analyzer/Dashboard "Export to Excel" produces an HTML table saved
with an `.xls` extension, not a real binary workbook. It works, but:

- No genuine Excel features (Table/autofilter, number formatting, multiple
  sheets) - it's HTML, so Excel opens it in compatibility mode.
- No structured export format at all for programmatic consumers (CSV is the
  closest thing, and it's untyped/unstructured for anything beyond a flat grid).
- No way to export a whole dashboard (all its pivot widgets) in one action -
  today that's one manual "Export to Excel" per widget, each producing its
  own separate HTML-as-.xls file.

## What we built to prove it's viable

This module adds, entirely from outside the product (subclassing/DOM-patching,
no product source changes):

1. Real `.xlsx` export (genuine Open XML, via IRIS Embedded Python + openpyxl)
   and JSON export, as new buttons in the Analyzer's existing Export dropdown
   and in each dashboard pivot widget's own Export dropdown.
2. A whole-dashboard export: one `.xlsx` workbook, one sheet per pivot widget,
   from one click in the dashboard viewer.
3. A REST API (`POST /xlsx`, `POST /json`) for pulling the same exports from
   scripts/automation without opening a browser session at all.
4. Availability detection: if Embedded Python/openpyxl isn't set up, every
   surface above degrades to a grayed-out button / `503` response with a
   plain-text reason, instead of a broken download.

Details: [README.md](../README.md) for the full user-facing feature set,
[Analyzer-export-classes.md](Analyzer-export-classes.md),
[Dashboard-export-classes.md](Dashboard-export-classes.md), and
[REST-API-classes.md](REST-API-classes.md) for exactly which classes implement
each piece.

## Recommendation: what to actually put in the JIRA ticket

Not everything above is a good product ask as-is - some of it is only
necessary *because* this had to be built outside the product. Splitting it:

### Ask for these (real product gaps)

1. **Real `.xlsx` export**, replacing (or added alongside, for compatibility)
   the current HTML-as-.xls "Export to Excel" - in both the Analyzer and
   dashboard pivot widgets. This is the core ask; everything else is smaller
   once this exists.
2. **JSON export** of a pivot's result set, next to the Excel option. Cheap
   for IS to add (no new dependency - a %DynamicObject already round-trips
   to JSON), and closes the "no structured format for scripts" gap without
   needing our REST workaround below.
3. **Whole-dashboard export** (one workbook, one sheet per widget), as a
   single dashboard-viewer action. If IS builds this server-side with a
   supported way to read each widget's *live* on-screen state (see the
   architecture note below), it wouldn't inherit this module's saved-state-only
   limitation.
4. **A supported REST endpoint** for triggering an export from automation,
   so customers don't need a bespoke `%CSP.REST` class per site. Lower
   priority than 1-3 - most of its value already exists via ad-hoc MDX/API
   scripting today.

### Don't carry these over as literal asks

These are things this module had to build only because the capability above
doesn't exist yet - once it does, these become unnecessary rather than
things IS should also implement:

- **Per-widget/per-page availability gray-out UI** (this module's `IsAvailable()`
  checks) - purely a workaround for an *optional* dependency (openpyxl) that
  a native implementation wouldn't have, since IS would presumably ship
  whatever it needs to generate `.xlsx` itself.
- **`DashboardTools` widget-type conversion helper** - a workaround for
  dashboards having no supported customization hook for a widget's export
  menu. Evidence that such a hook is missing, not something to ask IS to
  ship as-is.
- **DOM-patching the existing dropdown / small-menu overrides** - purely an
  outside-the-product technique (see the two class docs linked above for how
  much reverse-engineering that took); irrelevant once export is native.

### Worth a mention as a secondary/infra note

- **No supported server-side API to read a live (on-screen, unsaved) pivot
  widget's current state.** This is *why* our whole-dashboard export can only
  use each widget's saved definition - there's no browser session for a
  server classmethod to query. If IS's own multi-sheet export runs
  server-side the same way, it would hit the identical constraint unless
  paired with a client-side "gather all widgets' live state, then request
  the export" round trip. Worth flagging to engineering even if it doesn't
  become its own ticket.

## Suggested priority order for the ticket

1. Real `.xlsx` export (Analyzer + dashboard widgets)
2. JSON export
3. Whole-dashboard multi-sheet export
4. Supported export REST endpoint
