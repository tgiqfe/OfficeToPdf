import win32com.client
import pythoncom
import gc
import sys

input_file = sys.argv[1]
output_file = sys.argv[2]

pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Excel.Application")
app.Visible = False
workbooks = app.Workbooks
workbook = workbooks.Open(input_file, UpdateLinks=False, ReadOnly=True)
try:
    workbook.ExportAsFixedFormat(0, output_file, 0)
finally:
    workbook.Close(SaveChanges=False)
    app.Quit()
    del workbook
    del workbooks
    del app
    pythoncom.CoUninitialize()
    gc.collect()
