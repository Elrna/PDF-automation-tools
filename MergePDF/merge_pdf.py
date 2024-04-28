import tkinter as tk
from tkinterdnd2 import TkinterDnD
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
import os


def merge_pdfs(paths, output):
    pdf_writer = PdfWriter()
    for path in paths:
        try:
            pdf_reader = PdfReader(path)
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)
        except Exception as e:
            messagebox.showerror("PDF Read Error", f"Error reading {path}: {e}")
            return False

    try:
        with open(output, "wb") as out:
            pdf_writer.write(out)
    except IOError as e:
        messagebox.showerror("File Write Error", f"Error writing to {output}: {e}")
        return False

    return True


def get_dropfile_paths(event):
    files = root.tk.splitlist(event.data)
    pdf_paths = []
    for file in files:
        if os.path.isdir(file):
            for dirpath, dirs, filenames in os.walk(file):
                for name in filenames:
                    if name.lower().endswith(".pdf"):
                        pdf_paths.append(os.path.join(dirpath, name))
        elif file.lower().endswith(".pdf"):
            pdf_paths.append(file)
    pdf_paths.sort()
    return pdf_paths


def save_file_dialog(root):
    file_path = filedialog.asksaveasfilename(
        parent=root, defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]
    )
    if not file_path:
        messagebox.showinfo("No File Selected", "No output file was selected.")
        return None
    return file_path


def drop(event, root):
    pdf_paths = get_dropfile_paths(event)
    if pdf_paths:
        output = save_file_dialog(root)
        if output and merge_pdfs(pdf_paths, output):
            messagebox.showinfo(
                "Success", f"PDF files are successfully merged into {output}"
            )
        else:
            messagebox.showinfo("Cancelled", "PDF merge cancelled.")
    else:
        messagebox.showinfo("No PDFs", "No PDF files dropped.")
    return pdf_paths


# メイン処理
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.title("PDF Merger")

    label = tk.Label(root, text="Drop Files or Folder Here :D", width=60, height=25)
    label.pack(expand=True)

    label.drop_target_register("DND_Files")
    label.dnd_bind("<<Drop>>", lambda event, root=root: drop(event, root))

    root.mainloop()
