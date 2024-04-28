import os
import sys
import win32com.client as win32
import tkinter as tk

from tkinter import messagebox
from tkinterdnd2 import TkinterDnD
from threading import Thread

excel_app = None
word_app = None


def run_excel_macro(macro_name, target_file_path, convert_file_path):
    global excel_app
    try:
        if excel_app is None:
            excel_app = win32.Dispatch("Excel.Application")
            excel_app.Visible = False

        workbook = excel_app.Workbooks.Open(EXCEL_MACRO_PATH)

        excel_app.Application.Run(
            f"{workbook.Name}!{macro_name}", target_file_path, convert_file_path
        )

        workbook.Save()
        workbook.Close()
    except Exception as e:
        error_message = f"Failed to convert Excel file: {str(e)}"
        Thread(
            target=lambda message=error_message: messagebox.showerror("Error", message)
        ).start()


def run_word_macro(macro_name, target_file_path, convert_file_path):
    global word_app
    try:
        if word_app is None:
            word_app = win32.Dispatch("Word.Application")
            word_app.Visible = False

        document = word_app.Documents.Open(WORD_MACRO_PATH)

        word_app.Application.Run(macro_name, target_file_path, convert_file_path)

        document.Save()
        document.Close()
    except Exception as e:
        error_message = f"Failed to convert Word file: {str(e)}"
        Thread(
            target=lambda message=error_message: messagebox.showerror("Error", message)
        ).start()


def get_absolutePath(file_name):
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, file_name)


EXCEL_MACRO_PATH = get_absolutePath("ConvertPDF.xlsm")
WORD_MACRO_PATH = get_absolutePath("ConvertPDF.docm")


def convert_to_pdf(file_path):
    file_dir, file_name = os.path.split(file_path)
    new_file_path = os.path.join(
        file_dir,
        file_name.replace(".docx", ".pdf")
        .replace(".doc", ".pdf")
        .replace(".xlsx", ".pdf")
        .replace(".xls", ".pdf"),
    )

    if file_path.lower().endswith((".docx", ".doc")):
        run_word_macro("ConvertDocToPDF", file_path, new_file_path)
    elif file_path.lower().endswith((".xlsx", ".xls")):
        run_excel_macro("ConvertExcelToPDF", file_path, new_file_path)


def on_drop(event):
    file_paths = event.data.strip("{}").split("} {")
    for file_path in file_paths:
        convert_to_pdf(file_path.strip("{}"))


app = TkinterDnD.Tk()
app.title("PDF Converter")
app.geometry("500x300")

label = tk.Label(
    app, text="Drop Word or Excel files here:D\n(Extensions: docx, doc, xlsx, xls)"
)
label.pack(fill=tk.BOTH, expand=True)

app.drop_target_register("DND_Files")
app.dnd_bind("<<Drop>>", on_drop)


def close_office_apps():
    global excel_app, word_app
    if excel_app is not None:
        excel_app.Quit()
    if word_app is not None:
        word_app.Quit()
    app.destroy() 


app.protocol("WM_DELETE_WINDOW", close_office_apps)

app.mainloop()
