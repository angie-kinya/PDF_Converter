import pandas as pd
import os
import logging
from tabula import read_pdf  

# Configure logging
logging.basicConfig(filename="pdf_conversion.log", level=logging.INFO, 
                    format="%(asctime)s - %(levelname)s - %(message)s")

# Define source and output directories
PDF_DIR = "path_to_pdf_files"  # Change to actual PDF folder path
OUTPUT_DIR = "path_to_output_files"  # Change to actual output folder path
OUTPUT_FORMAT = "xlsx"  # Change to 'csv' if needed

def process_pdf_files():
    """Scans the PDF directory and converts each PDF file to the desired format."""
    try:
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)  # Ensure output directory exists
        
        pdf_files = [f for f in os.listdir(PDF_DIR) if f.lower().endswith(".pdf")]

        if not pdf_files:
            logging.info("No PDF files found for processing.")
            return

        for pdf_file in pdf_files:
            pdf_path = os.path.join(PDF_DIR, pdf_file)
            convert_pdf(pdf_path)

    except Exception as e:
        logging.error(f"Error processing PDF files: {str(e)}")

def convert_pdf(pdf_file):
    """Converts a single PDF file to an Excel or CSV file."""
    try:
        tables = read_pdf(pdf_file, pages="all", multiple_tables=True)  

        if not tables:
            logging.warning(f"No tables found in {pdf_file}. Skipping...")
            return

        combined_df = pd.concat(tables)
        output_file = os.path.join(OUTPUT_DIR, os.path.splitext(os.path.basename(pdf_file))[0] + "." + OUTPUT_FORMAT)

        if OUTPUT_FORMAT == "xlsx":
            combined_df.to_excel(output_file, index=False)
        elif OUTPUT_FORMAT == "csv":
            combined_df.to_csv(output_file, index=False)
        else:
            logging.error(f"Invalid output format: {OUTPUT_FORMAT}")
            return
        
        logging.info(f"Successfully converted {pdf_file} to {output_file}")

    except Exception as e:
        logging.error(f"Failed to convert {pdf_file}: {str(e)}")

# Run the script
if __name__ == "__main__":
    process_pdf_files()
