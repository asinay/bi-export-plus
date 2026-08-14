# Quick start: try it out in the Analyzer UI

Minimum steps to try the new **Excel Workbook**/**JSON** export buttons. (Skip
this if you need the standalone REST API too - see [README.md](README.md).)

## 1. Confirm Embedded Python is enabled, then install openpyxl

From an IRIS terminal, in any namespace:

```objectscript
Do ##class(%SYS.Python).Import("openpyxl")
```

- **`ModuleNotFoundError: No module named 'openpyxl'`** - Embedded Python
  itself is working, openpyxl just isn't installed yet. Install it into IRIS's
  embedded Python environment:

  ```bash
  <iris-install-dir>/bin/irispython -m pip install openpyxl
  ```

  Then re-run the `Do` command above - it should return silently (just the
  prompt, no error).

- **Any other error** (e.g. `<CLASS DOES NOT EXIST>`, or a message that Python
  isn't available/configured) - Embedded Python itself isn't enabled on this
  IRIS instance. That needs to be fixed at the instance level (by an IRIS
  admin, per your IRIS version's Embedded Python setup docs) before this
  module's Excel Workbook export can work.

## 2. Load the module

The simplest way for just trying out the UI: import the single packaged
classes file, [dist/BIExport-classes.xml](dist/BIExport-classes.xml), from an
IRIS terminal in your target namespace:

```objectscript
Do $system.OBJ.Load("C:\AI_Apps\bi-export-plus\dist\BIExport-classes.xml", "ck")
```

(No web application gets created this way, but none is needed for the
in-Analyzer export buttons below. See [README.md](README.md) for the IPM/zpm
install instead, which also sets up the REST API's web application.)

## 3. Open a pivot via the patched Analyzer page

Open a pivot in the normal Analyzer like you always do, then in the browser's
address bar replace `_DeepSee.UI.Analyzer.zen` with `BIExport.UI.Analyzer.cls`
and reload - everything else in the URL (namespace, `CUBE`, etc.) stays the
same.

## 4. Export

Open the pivot's **Export** dropdown - you'll see **Excel Workbook** and
**JSON** next to the stock Excel/CSV options. Pick one and it downloads.
