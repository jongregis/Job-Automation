import openpyxl as xl
from datetime import datetime

monthly_spreadsheet = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Invoices/2020 Enrollment/May 2020.xlsx"
pete_spreadsheet = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Pete's Backup/MILTARY/PETE ALL 3 SPREADSHEETS MYCAA FOR STACEY AND LISA/MAIN ENROLLMENT FOLDER/SPREADSHEETS/ECA ALL SCHOOLS MONTHLY SS.xlsx"


wb1 = xl.load_workbook(pete_spreadsheet)
wb2 = xl.load_workbook(monthly_spreadsheet)

broward = wb1.worksheets[1]
flagler = wb1.worksheets[3]

monthly = wb2.worksheets[2]


def findNextCell():
    for cell in monthly["C"]:
        if cell.value is None:

            return cell.row
            break
    else:
        return cell.row + 1


def findName(name):
    for cell in monthly["D"]:
        if cell.value == name:
            return True


def broward_students(current_month):
    mr = broward.max_row
    mc = broward.max_column
    num = findNextCell()

    for i in range(9, mr+1):

        c = broward.cell(row=i, column=3).value
        name = broward.cell(row=i, column=1).value
        last_number_row = num - 1
        if c == None:
            pass
        elif c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:

                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = broward.cell(row=i, column=3).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1
                monthly.cell(row=num, column=3).value = date1

                monthly.cell(row=num, column=4).value = name

                course = broward.cell(row=i, column=7).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                monthly.cell(row=num, column=9).value = 'BROWARD'
                monthly.cell(row=num, column=set_pricing_column(
                    'BROWARD')).value = broward.cell(row=i, column=9).value

                num += 1

    wb2.save(monthly_spreadsheet)


def flagler_students(current_month):
    mr = flagler.max_row
    mc = flagler.max_column
    num = findNextCell()

    for i in range(9, mr+1):

        c = flagler.cell(row=i, column=3).value
        name = flagler.cell(row=i, column=1).value
        last_number_row = num - 1
        if c == None:
            pass
        elif c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:

                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = flagler.cell(row=i, column=3).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1
                monthly.cell(row=num, column=3).value = date1

                monthly.cell(row=num, column=4).value = name

                course = flagler.cell(row=i, column=8).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                monthly.cell(row=num, column=9).value = 'FLAGLER'
                monthly.cell(row=num, column=set_pricing_column(
                    'FLAGLER')).value = flagler.cell(row=i, column=10).value

                num += 1

    wb2.save(monthly_spreadsheet)


def set_pricing_column(school):

    if school == "BROWARD":
        return 13
    elif school == "FLAGLER":
        return 12
    else:
        print("\033[1;31mno school with that name \033[0;0m")


def run_program_elearning():
    start = findNextCell()
    broward_students('05')
    flagler_students('05')
    wb2.save(monthly_spreadsheet)
    end = findNextCell()
    total = end-start
    print("Done transferring E-Learning Students")
    print("\033[1;32m{} \033[0;0mwere transferred".format(total))
