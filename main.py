import pandas as pd
import glob
from pathlib import Path
from fpdf import FPDF


filepath = glob.glob("excels/*.xlsx")

for path in filepath:
    df = pd.read_excel(path, sheet_name="Sheet 1")
    filename = Path(path).stem
    invoice_number = filename.split("-")[0]
    date = filename.split("-")[1]
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, f"Invoice Number: {invoice_number}", ln=1)
    pdf.set_font("Arial", size=8)
    pdf.cell(200, 10, f"Date: {date}", ln=1)

    col_widths = [35] * len(df.columns)

    for i, col in enumerate(df.columns):
        splitCol = col.split("_")
        name = " ".join(splitCol)
        final_name = name.title()
        pdf.set_font("Arial", size=8, style="B")
        pdf.cell(col_widths[i], 10, final_name, 1, 0, "C")
    pdf.ln()

    total_price = 0
    for i, row in df.iterrows():
        for j, col in enumerate(row):
            pdf.set_font("Arial", size=8)
            pdf.cell(col_widths[j], 10, f"{col}", 1, 0, "C")
            if j == len(row) - 1:
                total_price += col
        pdf.ln()
    pdf.set_font("Arial", size=8, style="B")
    for i, col in enumerate(df.columns):
        if col == "total_price":
            pdf.cell(col_widths[i], 10, f"{total_price}", 1, 0, "C")
        else:
            pdf.cell(col_widths[j], 10, "", 1, 0, "C")

    pdf.ln()
    pdf.ln()
    pdf.set_font("Arial", size=8, style="B")
    pdf.cell(200, 10, f"The total due amount is {total_price} Euros.")

    pdf.output(f"invoices/{invoice_number}.pdf")

