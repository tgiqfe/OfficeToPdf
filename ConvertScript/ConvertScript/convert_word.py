import win32com.client
import pythoncom
import gc
import sys

input_file = sys.argv[1]
output_file = sys.argv[2]

pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application")
app.Visible = False
doc = app.Documents.Open(input_file, ReadOnly=True)
try:
    doc.SaveAs(output_file, FileFormat=17)
finally:
    doc.Close(SaveChanges=False)
    app.Quit()
    del doc
    del app
    pythoncom.CoUninitialize()
    gc.collect()
