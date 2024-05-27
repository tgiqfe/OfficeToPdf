import win32com.client
import pythoncom
import gc
import sys
import os

input_file = sys.argv[1]
basename = os.path.basename(sys.argv[1])
basename_without_ext = os.path.splitext(basename)[0]
parent = os.path.dirname(sys.argv[1])
extension = os.path.splitext(basename)[1]
output_file = os.path.join(parent, basename_without_ext + ".pdf")

pythoncom.CoInitialize()
if extension == ".xls" or extension == ".xlsx":
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
elif extension == ".doc" or extension == ".docx":
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
elif extension == ".ppt" or extension == ".pptx":
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