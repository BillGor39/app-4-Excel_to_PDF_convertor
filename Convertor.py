from pathlib import Path
from fpdf import FPDF
import pandas as pd


def convert(filepath):
    # Get file name
    filename = Path(filepath).stem
    invoice_name, date = filename.split("-")

    # Create a PDF file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Make header
    pdf.set_font(family="Times", style="B", size=24)
    pdf.cell(w=0, h=12, txt=f"Invoice nr. {invoice_name}", border=0, ln=1)
    pdf.cell(w=0, h=12, txt=f"Date: {date}", border=0, ln=1)

    # Create table
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = [item.replace("_", " ").title() for item in df.columns]
    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=50, h=8, txt=columns[1], border=1)
    pdf.cell(w=50, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    pdf.set_font(family="Times", style="", size=12)
    for index, row in df.iterrows():
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    total = df["total_price"].sum()
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total), border=1, ln=1)

    pdf.cell(w=0, h=8, txt=f"The total price is {total}.", ln=1)

    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(w=24, h=8, txt="Edward")
    pdf.image("logo.png")

    pdf.output(f"PDFs_2/{filename}.pdf")


if __name__ == "__main__":
    convert("invoices/10001-2023.1.18.xlsx")
