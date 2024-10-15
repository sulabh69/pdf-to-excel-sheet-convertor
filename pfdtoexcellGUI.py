import pandas as pd

from tabula import read_pdf
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from tkinterdnd2 import DND_FILES, TkinterDnD

class PDFToExcelConverter:
    def __init__(self, master):
        self.master = master
        master.title("PDF to Excel Converter")

        self.label = tk.Label(master, text="Drag and drop a PDF file or browse:")
        self.label.pack(pady=10)

        self.pdf_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.entry = tk.Entry(master, textvariable=self.pdf_path, width=50)
        self.entry.pack(padx=10, pady=5)

        self.browse_button = tk.Button(master, text="Browse", command=self.browse_pdf)
        self.browse_button.pack(pady=5)

        self.output_entry = tk.Entry(master, textvariable=self.output_path, width=50)
        self.output_entry.pack(padx=10, pady=5)

        self.output_button = tk.Button(master, text="Set Output Location", command=self.set_output_location)
        self.output_button.pack(pady=5)

        self.convert_button = tk.Button(master, text="Convert to Excel", command=self.convert_pdf_to_excel)
        self.convert_button.pack(pady=20)

        # Drag and drop
        master.drop_target_register(DND_FILES)
        master.dnd_bind('<<Drop>>', self.drop)

    def browse_pdf(self):
        file_path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.pdf_path.set(file_path)

    def set_output_location(self):
        file_path = filedialog.asksaveasfilename(title="Save Excel File", defaultextension=".xlsx", 
                                                   filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.output_path.set(file_path)

    def convert_pdf_to_excel(self):
        pdf_file = self.pdf_path.get()
        excel_file = self.output_path.get()

        if not pdf_file or not excel_file:
            messagebox.showerror("Error", "Please select a PDF file and set the output location.")
            return

        try:
            tables = read_pdf(pdf_file, pages='all', multiple_tables=True)
            if tables:
                with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                    for i, table in enumerate(tables):
                        table.to_excel(writer, sheet_name=f'Table {i + 1}', index=False)
                messagebox.showinfo("Success", f"Successfully converted '{pdf_file}' to '{excel_file}'")
            else:
                messagebox.showwarning("Warning", "No tables found in the PDF.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def drop(self, event):
        pdf_file = event.data.strip('{}')  # Remove curly braces
        if pdf_file.lower().endswith('.pdf') and os.path.isfile(pdf_file):
            self.pdf_path.set(pdf_file)

if __name__ == "__main__":
    root = TkinterDnD.Tk()  # Use TkinterDnD instead of tk.Tk()
    app = PDFToExcelConverter(root)
    root.mainloop()
