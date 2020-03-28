from .sampleInvoice import create_invoice
import openpyxl as xl


mycaa_invoice = "/Users/jongregis/Python/JobAutomation/JobAutomation/MYCAA Automation.xlsm"
wb2 = xl.load_workbook(filename=mycaa_invoice, read_only=False, keep_vba=True)

data_sheet = wb2.worksheets[1]


def excel_to_pdf():
    mr = data_sheet.max_row
    for i in range(2, mr+1):
        start_date = data_sheet.cell(row=i, column=1).value
        name = data_sheet.cell(row=i, column=2).value
        description = data_sheet.cell(row=i, column=3).value
        school = data_sheet.cell(row=i, column=4).value
        invoice_number = data_sheet.cell(row=i, column=5).value
        total = data_sheet.cell(row=i, column=6).value
        create_invoice(start_date, name, description,
                       total, '', school, invoice_number)

    print("PDF Inovices Done!")