from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

file_paths = glob.glob("invoices/*.xlsx")
for filepath in file_paths:
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

    # Read the file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add header
    columns = [item.replace("_", " "). title() for item in list(df.columns)]
    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(w=30, h=6, txt=str(columns[0]), border=1)
    pdf.cell(w=50, h=6, txt=str(columns[1]), border=1)
    pdf.cell(w=50, h=6, txt=str(columns[2]), border=1)
    pdf.cell(w=35, h=6, txt=str(columns[3]), border=1)
    pdf.cell(w=30, h=6, txt=str(columns[4]), ln=1, border=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", style="", size=12)
        pdf.cell(w=30, h=6, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=6, txt=str(row["product_name"]), border=1)
        pdf.cell(w=50, h=6, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=35, h=6, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=6, txt=str(row["total_price"]), ln=1, border=1)

    # Add sum of total price
    total_price_sum = df["total_price"].sum()
    pdf.set_font(family="Times", style="", size=12)
    pdf.cell(w=30, h=6, txt="", border=1)
    pdf.cell(w=50, h=6, txt="", border=1)
    pdf.cell(w=50, h=6, txt="", border=1)
    pdf.cell(w=35, h=6, txt="", border=1)
    pdf.cell(w=30, h=6, txt=str(total_price_sum), ln=1, border=1)

    # Add total sum sentence
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=8, txt=f"The total price is {total_price_sum}", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=24, h=8, txt="Edward")
    pdf.image("logo.png")

    pdf.output(f"PDFs/{filename}.pdf")


