import win32com.client
import pythoncom
import gc
import os
import sys

def convert_powerpoint_to_pdf(input_file, output_file):
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

basename = os.path.basename(sys.argv[1])
basename_without_ext = os.path.splitext(basename)[0]
dirname = os.path.dirname(sys.argv[1])
output_file = os.path.join(dirname, basename_without_ext + ".pdf")

print(sys.argv[1])
print(output_file)

# PDF化実行
convert_powerpoint_to_pdf(sys.argv[1], output_file)
