from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

file_paths = glob.glob("invoices/*.xlsx")
for filepath in file_paths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_name = filename.split("-")[0]
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=8, txt=f"Invoice nr. {invoice_name}", border=0)
    pdf.output(f"PDFs/{filename}.pdf")


