import win32com.client
import pythoncom
import gc
import sys

input_file = sys.argv[1]
output_file = sys.argv[2]

pythoncom.CoInitialize()
app = win32com.client.Dispatch("Powerpoint.Application")
doc = app.Presentations.Open(input_file, ReadOnly=True, WithWindow=False)
try:
    doc.SaveAs(output_file, FileFormat=32)
finally:
    doc.Close()
    app.Quit()
    del doc
    del app
    pythoncom.CoUninitialize()
    gc.collect()

