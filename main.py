from pathlib import Path
from borb.pdf import Document
from borb.pdf.page.page import Page
from borb.pdf.canvas.layout.page_layout.multi_column_layout import SingleColumnLayout
from decimal import Decimal
from borb.pdf.canvas.layout.image.image import Image
from borb.pdf.canvas.layout.text.paragraph import Paragraph
from borb.pdf.canvas.layout.layout_element import Alignment
from datetime import datetime
from borb.pdf.pdf import PDF
from borb.pdf.canvas.layout.table.fixed_column_width_table import FixedColumnWidthTable as Table
from borb.pdf.canvas.layout.table.table import TableCell
from borb.pdf.canvas.color.color import HexColor, X11Color
import tkinter
from tkinter import filedialog
from tkinter import messagebox
window = tkinter.Tk()
window.withdraw()
Company_name = input("Whats your company's name? ")
INVOICE_NUM = input("Please Enter the invoice number: ")
our_addr1 = input("Please Enter your 1st line address: ")
our_addr2 = input("Please Enter your 2nd line address: ")
our_country = input("Please Enter your company's country: ")
address_1 = input("Please Enter the Client's first address row: ")
address_2 = input("Please Enter the Client's Second address Row: ")
Country = input("Please enter the country: ")
CURSYMB = input("Please enter the Symbol of currency: ")
CURSYMB = CURSYMB + " "
PROD_AMT = input("Please enter how many products are there: ")
Array = [[input("Name Of product: "), input("QTY: "), input("Unit Price: ")] for _ in range(int(PROD_AMT))]
logo = filedialog.askopenfilename(
    title="Choose your logo",
    filetypes=[('image files', '.png'),
               ('image files', '.jpg')
               ]
)
def _build_top_part():
    table_002 = Table(number_of_rows=1,number_of_columns=2)
    table_002.add(
        Image(
        Path(logo),
        width=Decimal(128),
        height=Decimal(128),
        ))
    table_002.add(
        Paragraph(" ")
    )
    table_002.set_padding_on_all_cells(Decimal(2),Decimal(2),Decimal(2), Decimal(2))
    table_002.no_borders()
    return table_002
def _build_company_info():
    table_003 = Table(number_of_rows=4,number_of_columns=1)
    table_003.add(
        Paragraph(
            Company_name,
            font_size=Decimal(20),
            horizontal_alignment=Alignment.LEFT,
            font="Helvetica-bold"
        )
    )
    table_003.add(
        Paragraph(
            our_addr1,
            font_size=Decimal(15),
            horizontal_alignment=Alignment.LEFT,
            font="Helvetica-bold"
        )
    )
    table_003.add(
        Paragraph(
            our_addr2,
            font_size=Decimal(15),
            horizontal_alignment=Alignment.LEFT,
            font="Helvetica-bold"
        )
    )
    table_003.add(
        Paragraph(
            our_country,
            font_size=Decimal(15),
            horizontal_alignment=Alignment.LEFT,
            font="Helvetica-bold"
        )
    )
    table_003.set_padding_on_all_cells(Decimal(2),Decimal(2),Decimal(2), Decimal(2))
    table_003.no_borders()
    return table_003
def _build_billing_and_shipping_information():
    table_001 = Table(number_of_rows=4, number_of_columns=3)
    table_001.add(
        Paragraph(
            "BILL TO",
          font="Helvetica-bold"
        )
    )
    table_001.add(Paragraph("  "))
    table_001.add(
        Paragraph(
            " "
        )
    )
    table_001.add(Paragraph(address_1))        # BILLING
    table_001.add(Paragraph("Date", font="Helvetica-Bold", horizontal_alignment=Alignment.RIGHT))        # SHIPPING
    now = datetime.now()
    table_001.add(Paragraph("%d/%d/%d" % (now.day, now.month, now.year)))

    table_001.add(Paragraph(address_2))          # BILLING
    table_001.add(Paragraph("Invoice #", font="Helvetica-Bold", horizontal_alignment=Alignment.RIGHT))
    table_001.add(Paragraph(INVOICE_NUM))

    table_001.add(Paragraph(Country))        # BILLING
    table_001.add(Paragraph("Due Date", font="Helvetica-Bold", horizontal_alignment=Alignment.RIGHT))
    table_001.add(Paragraph("%d/%d/%d" % (now.day, now.month, now.year)))

    table_001.set_padding_on_all_cells(Decimal(2), Decimal(2), Decimal(2), Decimal(2))
    table_001.no_borders()
    return table_001


def _build_itemized_description_table():
    #3 + int(PROD_AMT)
    num_row = 3 + int(PROD_AMT) - 1
    table_001 = Table(number_of_rows=num_row, number_of_columns=4)
    for h in ["DESCRIPTION", "QTY", "UNIT PRICE", "AMOUNT"]:
        table_001.add(
            TableCell(
                Paragraph(h, font_color=X11Color("White")),
                background_color=HexColor("#34495e"),
            )
        )

    odd_color = HexColor("FFFFFF")
    even_color = HexColor("FFFFFF")
    Total_Price = 0
    for row_number, item in enumerate(Array):
        c = even_color if row_number % 2 == 0 else odd_color
        table_001.add(TableCell(Paragraph(item[0]), background_color=c))
        table_001.add(TableCell(Paragraph(str(item[1])), background_color=c))
        table_001.add(TableCell(Paragraph(CURSYMB + str(item[2])), background_color=c))
        table_001.add(TableCell(Paragraph(CURSYMB + str(int(item[1]) * int(item[2])), horizontal_alignment=Alignment.RIGHT), background_color=c))
        Total_Price += int(item[1]) * int(item[2])
    Final_Price = CURSYMB + str(Total_Price)
    table_001.add(
        TableCell(Paragraph("Total", font="Helvetica-Bold", horizontal_alignment=Alignment.RIGHT), col_span=3, ))
    table_001.add(TableCell(Paragraph(Final_Price, horizontal_alignment=Alignment.RIGHT)))
    table_001.set_padding_on_all_cells(Decimal(2), Decimal(2), Decimal(2), Decimal(0))
    table_001.no_borders()
    return table_001
# Create document
pdf = Document()
# Add page
page = Page()
pdf.append_page(page)
page_layout = SingleColumnLayout(page)
page_layout.vertical_margin = page.get_page_info().get_height() * Decimal(0.02)
page_layout.add(_build_top_part())
page_layout.add(_build_company_info())
page_layout.add(Paragraph(" "))
# Invoice information table
page_layout.add(_build_billing_and_shipping_information())
# Itemized description
page_layout.add(_build_itemized_description_table())
page_layout.add((Paragraph("THANK YOU FOR YOUR BUSINESS!",font="Helvetica-bold",font_color=HexColor("#3498db"),horizontal_alignment=Alignment.CENTERED)))
#page_layout.add(_build_additional_info())
with open("invoice.pdf", "wb") as pdf_file_handle:
    PDF.dumps(pdf_file_handle, pdf)
tkinter.messagebox.showinfo(title="Invoice Generator", message="Exported as invoice.pdf")