from fpdf import FPDF
import pandas as pd
import glob

file_paths = glob.glob("invoices/*.xlsx")
for file in file_paths:
    df = pd.read_excel(file, sheet_name="Sheet 1")
    print(df)
