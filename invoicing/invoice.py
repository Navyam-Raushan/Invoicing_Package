import glob
import os
import pandas as pd
from fpdf import FPDF
from pathlib import Path


def generate(invoices_path, output_dir_name, image_path, company_name, product_id, product_name,
             amount_purchased, price_per_unit, total_price):

    """
    This function will convert invoice excel files into a pdf file.
    Input the required parameters here, Also fill the column names
    your Excel file as parameter.
    :param company_name:
    :param invoices_path:
    :param output_dir_name:
    :param image_path:
    :param product_id:
    :param product_name:
    :param amount_purchased:
    :param price_per_unit:
    :param total_price:
    :return:
    """
    
    # Firstly reading the excels file, glob help to read similiar files.
    filepaths = glob.glob(f"{invoices_path}/*.xlsx")

    # for getting all data in these xl sheets
    # for processing xl files we need openpyxl library of python.
    for filepath in filepaths:

        # Extracting names of files to use it, firstline will give pure name by removing suffix
        """It will return filename 10001-2023.01.08.xlsx 
        like this.
        """
        filename = Path(filepath).stem
        invoice_nr, date = filename.split("-")

        # now generating pdf files.
        pdf = FPDF(orientation="l", unit="mm", format="a4")
        pdf.add_page()

        pdf.set_font(family="Times", size=18, style="B")
        pdf.cell(w=50, h=8, ln=1, txt=f"Invoice nr: {invoice_nr}")
        pdf.cell(w=50, h=2, ln=1, txt="")

        pdf.cell(w=50, h=8, ln=1, txt=f"Date: {date}")
        pdf.cell(w=50, h=8, ln=1, txt="")

        # Reading df and columns
        df = pd.read_excel(filepath, sheet_name="Sheet 1")
        columns = list(df.columns)
        columns = [item.replace("_", " ").title() for item in columns]

        # ADDING HEADER
        pdf.set_font(family="Times", size=12, style="B")
        pdf.cell(h=8, w=30, txt=columns[0], align="l", border=1)
        pdf.cell(h=8, w=70, txt=columns[1], align="l", border=1)
        pdf.cell(h=8, w=50, txt=columns[2], align="r", border=1)
        pdf.cell(h=8, w=30, txt=columns[3], align="r", border=1)
        # At last cell we must add ln=1 to get next value
        pdf.cell(h=8, w=30, txt=columns[4], align="r", border=1, ln=1)

        # ADDING CELLS TO TABLE
        for index, row in df.iterrows():
            pdf.set_font(family="Times", size=10)
            pdf.set_text_color(80, 80, 80)

            # ADDING CELLS TO TABLE
            # here txt must be string not integer.
            pdf.cell(h=8, w=30, txt=f"{row[product_id]}", align="l", border=1)
            pdf.cell(h=8, w=70, txt=row[product_name], align="l", border=1)
            pdf.cell(h=8, w=50, txt=f"{row[amount_purchased]}", align="r", border=1)
            pdf.cell(h=8, w=30, txt=f"{row[price_per_unit]}", align="r", border=1)
            # At last cell we must add ln=1 to get next value
            pdf.cell(h=8, w=30, txt=f"{row[total_price]}", align="r", border=1, ln=1)

        # ADDING TOTAL SUM
        total_sum = df[total_price].sum()
        pdf.set_font(family="Times", size=12, style="B")
        pdf.cell(h=8, w=30, txt="Final Amount", align="l", border=1)
        pdf.cell(h=8, w=70, txt=" ", align="l", border=1)
        pdf.cell(h=8, w=50, txt=" ", align="r", border=1)
        pdf.cell(h=8, w=30, txt=" ", align="r", border=1)
        # At last cell we must add ln=1 to get next value
        pdf.cell(h=8, w=30, txt=str(total_sum), align="r", border=1, ln=1)

        # Output lines
        pdf.set_font(family="Times", size=14, style="B")
        pdf.set_text_color(0, 0, 0)
        pdf.cell(h=5, w=0, txt=" ", ln=1)
        pdf.cell(h=8, w=0, txt=f"The amount due is RS {total_sum}.", ln=1)
        pdf.cell(h=2, w=0, txt=" ", ln=1)

        # Company name and logo
        pdf.cell(h=8, w=26, txt=f"{company_name}", align="l")
        pdf.image(f"{image_path}.png", w=10)

        # Create a pdf directory before output (If not present)
        if not os.path.exists(output_dir_name):
            os.makedirs(output_dir_name)
        # Must be outside from nested loop.
        pdf.output(f"{output_dir_name}/{filename}.pdf")
