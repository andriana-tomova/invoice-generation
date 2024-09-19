import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
pdf = FPDF(orientation="P", unit="mm", format="A4" )

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(100, 100, 100)

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    pdf.cell(w=30, h=10, txt=f"Invoice Nr. {invoice_nr}", align="L", ln=1)

    pdf.set_font(family="Times", style="I", size=8)
    pdf.cell(w=30, h=5, txt=f"Date {date}", align="L", ln=1)

    # Excel data goes to dataframe object
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add header of the table
    columns = list(df.columns)
    columns = [item.replace("_", " ").capitalize() for item in columns]
    pdf.set_font(family="Times", style="B", size=8)
    pdf.cell(w=30, h=8, txt=str(columns[0]), border=1)
    pdf.cell(w=40, h=8, txt=str(columns[1]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[4]), border=1, ln=1)

    # Add rows of the tablw
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["product_name"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Add total sum
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=8)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add company name and logo
    pdf.cell(w=100, h=8, txt="", ln=1)
    pdf.cell(w=15, h=8, txt=f"PythonHow")
    pdf.image("logo.png", w=8)


    pdf.output(f"PDFs/{filename}.pdf")








