from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

file_paths = glob.glob("invoices/*.xlsx")
for filepath in file_paths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Create a pdf file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Get the file name and date
    filename = Path(filepath).stem
    invoice_name = filename.split("-")[0]
    date = filename.split("-")[1]

    # Write on the pdf
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=8, txt=f"Invoice nr. {invoice_name}", border=0, ln=1)
    pdf.cell(w=0, h=8, txt=f"Date: {date}", border=0, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")


