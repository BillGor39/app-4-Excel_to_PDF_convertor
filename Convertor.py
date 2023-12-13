from pathlib import Path
from fpdf import FPDF
import pandas as pd

def convert(filepath):
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    filename = Path(filepath).stem
    invoice_name = filename.split("-")[0]

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=24)
    pdf.cell(w=0, h=12, txt=f"Invoice nr. {invoice_name}", border=0, align="L")
    pdf.output(f"PDFs_2/{invoice_name}.pdf")