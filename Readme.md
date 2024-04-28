# PDF Automation Tools

This repository includes a collection of tools for converting, merging, and manipulating PDF files using Python. Each tool is designed with a graphical user interface (GUI) to simplify user interactions. Below is a description of each script and its functionality.

## Tools Included

## 1. PDF Converter (convert_pdf.py)
This script provides a GUI for converting Microsoft Word and Excel documents to PDF. Users can drag and drop files into the window to convert them automatically. The script utilizes macros embedded in Word and Excel documents (ConvertPDF.docm and ConvertPDF.xlsm) to perform the conversion.

### Features:
- Supports .docx, .doc, .xlsx, and .xls files.
- Drag-and-drop functionality for easy file conversion.
- Automatic handling of conversion errors with user alerts.

## 2. PDF Merger (merge_pdf.py)
Merge multiple PDF files into a single PDF document. The tool supports dragging and dropping files or folders containing PDFs, which are then merged into a specified output file.

### Features:
- Drag-and-drop interface for files and folders.
- Sorts and merges all PDF files found in the dropped items.
- Error handling for file reading and writing issues.

## 3. PDF Page Remover (delete_pdfpage.py)
Allows users to remove specific pages from a PDF file. The tool accepts a PDF file dropped into its interface, then prompts the user to specify the pages to be removed, either as a list of numbers or a range.

### Features:
- Remove specified pages from a PDF document.
- GUI prompts for selecting pages to remove.
- Option to save the modified PDF with selected pages removed.

## Installation
To run these scripts, you will need Python installed on your system along with the following packages:

- "tkinterdnd2"
- "PyPDF2"
- "win32com.client"
- "tkinter"

You can install the required packages using pip:
pip install pypiwin32 PyPDF2 tkinterdnd2

## Usage
Each script can be executed individually by running the Python file. Ensure that the necessary macro-enabled documents (ConvertPDF.xlsm and ConvertPDF.docm) are present in the same directory as the scripts for the PDF Converter.

python convert_pdf.py  # For converting documents to PDF
python merge_pdf.py    # For merging PDF documents
python delete_pdfpage.py  # For removing pages from a PDF

## Note
These tools are designed to work on Windows due to the dependency on win32com.client for Microsoft Office automation.