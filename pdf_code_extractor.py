import fitz  # PyMuPDF
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os

def extract_text_from_pdf(pdf_path):
    document = fitz.open(pdf_path)
    text = ""
    for page_num in range(document.page_count):
        page = document.load_page(page_num)
        text += page.get_text()
    return text

def find_codes_in_text(text):
    pattern = r'COD\.\s*(\d{6})'
    codes = re.findall(pattern, text)
    return codes

def create_excel_from_codes(codes, excel_path):
    df = pd.DataFrame(codes, columns=["Codes"])
    df.to_excel(excel_path, index=False)

def select_pdf_and_extract_codes():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        text = extract_text_from_pdf(pdf_path)
        codes = find_codes_in_text(text)
        if codes:
            excel_filename = "extracted_codes.xlsx"
            excel_path = os.path.join(os.path.dirname(pdf_path), excel_filename)
            create_excel_from_codes(codes, excel_path)
            messagebox.showinfo("Processed", "Codes have been processed.")
        else:
            messagebox.showinfo("No Codes Found", "No 6-digit codes were found in the selected PDF.")

# GUI setup
root = tk.Tk()
root.title("PDF Code Extractor")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack(padx=10, pady=10)

button = tk.Button(frame, text="Select PDF", command=select_pdf_and_extract_codes)
button.pack()

root.mainloop()
