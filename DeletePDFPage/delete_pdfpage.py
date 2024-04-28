import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, Checkbutton, IntVar
from tkinterdnd2 import TkinterDnD
from PyPDF2 import PdfReader, PdfWriter

def parse_page_numbers(page_str):
    """ページ指定の文字列を解析してページ番号のリストを返す"""
    page_nums = []
    parts = page_str.split(',')
    for part in parts:
        if '-' in part:
            start, end = map(int, part.split('-'))
            page_nums.extend(range(start, end + 1))
        else:
            page_nums.append(int(part))
    return list(set(page_nums))


def remove_pages(pdf_path, pages_to_remove):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    num_pages = len(reader.pages)
    pages_to_remove = [p - 1 for p in pages_to_remove if 0 < p <= num_pages]
    pages_to_keep = [p for p in range(num_pages) if p not in pages_to_remove]
    for page_number in pages_to_keep:
        writer.add_page(reader.pages[page_number])
    return writer, reader.pages


def save_pdf(writer, output_path):
    if output_path:
        with open(output_path, 'wb') as f:
            writer.write(f)
        messagebox.showinfo("Success", "PDF saved successfully.")


def handle_drop(event, split_var):
    file_path = event.data.strip("{}")
    if not file_path.lower().endswith(".pdf"):
        messagebox.showinfo("Invalid File", "Please drop a PDF file.")
        return

    pages = simpledialog.askstring("Remove Pages", "Enter pages to remove (comma-separated or range):")
    if pages:
        try:
            pages_to_remove = parse_page_numbers(pages)
            writer, reader_pages = remove_pages(file_path, pages_to_remove)
            output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            save_pdf(writer, output_path)
        except ValueError:
            messagebox.showerror("Error", "Invalid page number format.")
    else:
        messagebox.showinfo("No Pages", "No pages were removed.")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.title("PDF Page Remover")
    label = tk.Label(root, text="Drop PDF Here:D", width=60, height=25)
    label.pack(expand=True)
    label.drop_target_register("DND_Files")
    split_var = IntVar(value=0)
    split_checkbox = Checkbutton(root, text="Save as separate files", variable=split_var)
    split_checkbox.pack()
    label.dnd_bind("<<Drop>>", lambda event, sv=split_var: handle_drop(event, sv))
    root.mainloop()
