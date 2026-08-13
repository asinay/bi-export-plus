Class BIExport.PyTest Extends %RegisteredObject
{

ClassMethod BuildTiny() As %Binary [ Language = python ]
{
import openpyxl
import io

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Test"
ws["A1"] = "Hello"
ws["B1"] = 42
ws["B1"].number_format = "#,##0"

buf = io.BytesIO()
wb.save(buf)
return buf.getvalue()
}

ClassMethod WriteTinyToFile(pFile As %String) As %Status
{
    Set tSC = $$$OK
    Try {
        Set tBytes = ..BuildTiny()
        Set tStream = ##class(%Stream.FileBinary).%New()
        Set tSC = tStream.LinkToFile(pFile)
        If $$$ISERR(tSC) Quit
        Do tStream.Write(tBytes)
        Set tSC = tStream.%Save()
    }
    Catch ex {
        Set tSC = ex.AsStatus()
    }
    Quit tSC
}

ClassMethod BraceTest(pJSON As %String) As %String [ Language = python ]
{
import json
data = json.loads(pJSON)
d = {"a": 1, "b": {"nested": [1, 2, 3]}}
name = data.get("name", "world")
msg = f"hello {name}, keys={list(d.keys())}"
return msg
}

}
