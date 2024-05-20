import win32com.client
import pythoncom
import gc
import os
import sys

def convert_excel_to_pdf(input_file, output_file):
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

basename = os.path.basename(sys.argv[1])
basename_without_ext = os.path.splitext(basename)[0]
dirname = os.path.dirname(sys.argv[1])
output_file = os.path.join(dirname, basename_without_ext + ".pdf")

print(sys.argv[1])
print(output_file)

convert_excel_to_pdf(sys.argv[1], output_file)
