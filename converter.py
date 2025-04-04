import pandas as pd
import pdfplumber
import os
import logging
from tabula import read_pdf  
import tabula
import tkinter as tk
from tkinter import filedialog, messagebox

# Configure logging
logging.basicConfig(filename="pdf_conversion.log", level=logging.INFO, 
                    format="%(asctime)s - %(levelname)s - %(message)s")

# Define a class that holds the GUI for the conversion
class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel/CSV Converter")
        self.root.geometry("400x200")
        self.root.resizable(False, False)

        # Label - File Selection
        self.label = tk.Label(root, text="Select a PDF file to convert:")
        self.label.pack(pady=10)

        # Button - Choose file
        self.select_button = tk.Button(root, text="Choose PDF file", command=self.select_pdf_file)
        self.select_button.pack(pady=5)

        # Dropdown - Choose output format
        self.format_var = tk.StringVar(root)
        self.format_var.set("xlsx") # Default to microsoft excel format
        self.format_label = tk.Label(root, text="Select output format:")
        self.format_label.pack(pady=5)
        self.format_dropdown = tk.OptionMenu(root, self.format_var, "xlsx", "csv")
        self.format_dropdown.pack(pady=5)

        # Button - Covert the pdf
        self.convert_button = tk.Button(root, text="Convert", command=self.convert_pdf)
        self.convert_button.pack(pady=10)

        # File path storage
        self.pdf_file = None

    # Method to select the pdf file
    def select_pdf_file(self):
        """Opens a file dialog for the user to select a PDF file."""
        file_path = filedialog.askopenfilename(title="Select a PDF file", filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_file = file_path
            messagebox.showinfo("File selected", f"Selected file:\n{self.pdf_file}")

    # Method for conversion
    def convert_pdf(self):
        """Converts the selected PDF file to the chosen format."""
        if not self.pdf_file:
            messagebox.showwarning("No file selected", "Please select a PDF file first.")
            return

        output_format = self.format_var.get()
        try:
            all_tables = []

            with pdfplumber.open(self.pdf_file) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        all_tables.append(df)

            if not all_tables:
                messagebox.showwarning("No Tables Found", "No tabular data found in the PDF.")
                return

            final_df = pd.concat(all_tables)
            output_file = os.path.splitext(self.pdf_file)[0] + "." + output_format

            if final_df.index.duplicated().any():
                final_df = final_df.reset_index(drop=True)
                logging.warning("Duplicate index found. Resetting index.")
                messagebox.showinfo("Duplicate Index", "Duplicate index found. Resetting index.")
                
            if output_format == "xlsx":
                final_df.to_excel(output_file, index=False)
            elif output_format == "csv":
                final_df.to_csv(output_file, index=False)

            messagebox.showinfo("Conversion Successful", f"File saved as:\n{output_file}")

        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed:\n{str(e)}")

# Main function to run the GUI
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()