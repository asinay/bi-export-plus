# Classes needed for the standalone REST API

The standalone `/xlsx` and `/json` REST endpoints (see README's "Standalone
REST API" section for the request/response contract) are implemented across
these classes:

- **`BIExport.REST`** - the `%CSP.REST` dispatch class. Its `UrlMap` maps
  `POST /xlsx` to `ExportXLSX()` and `POST /json` to `ExportJSON()`. Both
  parse the request body's `mdx`/`filename`/`title`/... fields into a
  `%DynamicObject` and pass it straight through to `MDXExport`; neither
  route does any MDX/export logic itself. Also owns `%DefaultFileName()`,
  the fallback filename used when the caller doesn't supply one.

- **`BIExport.MDXExport`** - shared with the Analyzer/dashboard export paths
  (see [Analyzer-export-classes.md](Analyzer-export-classes.md)). Runs the
  MDX and shapes the result into the row/column/data structure both routes
  return or hand to the writer below (`BuildExportData()`).

- **`BIExport.XLSXWriter`** - `/xlsx` only. Turns `MDXExport`'s data into the
  actual `.xlsx` bytes (`BuildXLSX()`), and its `IsAvailable()` check is what
  makes `ExportXLSX()` return `503` instead of a broken download when IRIS
  Embedded Python/openpyxl aren't set up. `/json` doesn't call this class at
  all - `ExportJSON()` returns `MDXExport`'s output via `%ToJSON()` directly,
  no Python involved.

This API also needs a `%CSP.REST`-dispatched web application pointed at
`BIExport.REST` (`/csp/bi-export` by default) - IPM creates this from
[module.xml](../module.xml) automatically; manual installs create it by hand
(README's "Manual install" section, step 2). The in-Analyzer export buttons
and dashboard export don't need this web application at all - they're plain
CSP pages in the Analyzer's/dashboard's own application, not REST calls.

## Request flow

```
POST /csp/bi-export/xlsx or /json
  -> BIExport.REST.ExportXLSX() / ExportJSON()
       -> BIExport.XLSXWriter.IsAvailable()   (xlsx only, 503 short-circuit)
       -> BIExport.MDXExport.BuildExportData()
       -> BIExport.XLSXWriter.BuildXLSX()      (xlsx only)
```
