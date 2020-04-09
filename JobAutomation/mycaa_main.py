import openpyxl as xl
from datetime import datetime
import os

jon_email_workbook = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Letters to students/Weekly Email for Lisa/Jon weekly email list.xlsx"
pete_spreadsheet = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Pete's Backup/MILTARY/PETE ALL 3 SPREADSHEETS MYCAA FOR STACEY AND LISA/MAIN ENROLLMENT FOLDER/SPREADSHEETS/students mycaa FINAL.xlsx"
monthly_spreadsheet = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Invoices/2020 Enrollment/April 2020.xlsx"
# assert os.path.exists(pete_spreadsheet)

lastMonth = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Invoices/2020 Enrollment/March 2020.xlsx"
# fileName = 'test1.xlsx'
wb1 = xl.load_workbook(pete_spreadsheet)
auburn = wb1.worksheets[0]
clemson = wb1.worksheets[1]
csu = wb1.worksheets[2]
lsu = wb1.worksheets[3]
msu = wb1.worksheets[4]
unh = wb1.worksheets[5]
tamu = wb1.worksheets[7]
wku = wb1.worksheets[8]
uwlax = wb1.worksheets[9]
desu = wb1.worksheets[10]

fileName1 = 'test2.xlsx'
wb2 = xl.load_workbook(monthly_spreadsheet)
monthly = wb2.worksheets[0]

wb3 = xl.load_workbook(jon_email_workbook)
jon_sheet = wb3.active


def findNextCell():
    for cell in monthly["C"]:
        if cell.value is None:
            return cell.row
            break
    else:
        return cell.row + 1


def findNextCellJonEmail():
    for cell in jon_sheet["A"]:

        if cell.value is None:
            return cell.row
            break
    else:
        return cell.row + 1


def findName(name):
    for cell in monthly["D"]:
        if cell.value == name:
            return True


def auburn_students(current_month):
    mr = auburn.max_row
    mc = auburn.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(5961, mr+1):

        c = auburn.cell(row=i, column=3).value
        name = auburn.cell(row=i, column=9).value
        last_number_row = num - 1
        if c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:

                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = auburn.cell(row=i, column=1).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1

                date2 = auburn.cell(row=i, column=3).value
                date2 = date2.strftime('%m') + '/' + \
                    date2.strftime('%d') + '/' + date2.strftime('%-y')
                monthly.cell(row=num, column=3).value = date2

                date3 = auburn.cell(row=i, column=4).value
                date3 = date3.strftime('%m') + '/' + \
                    date3.strftime('%d') + '/' + date3.strftime('%-y')
                jon_sheet.cell(row=num1, column=3).value = date3

                address = auburn.cell(row=i, column=7).value
                monthly.cell(row=num, column=14).value = address

                email = auburn.cell(row=i, column=8).value
                jon_sheet.cell(row=num1, column=2).value = email

                if 'LAPTOP' in name:
                    monthly.cell(row=num, column=8).value = 'x'
                monthly.cell(row=num, column=4).value = name
                jon_sheet.cell(row=num1, column=1).value = name

                course = auburn.cell(row=i, column=10).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                # code = auburn.cell(row=i, column=11).value
                # monthly.cell(row=num, column=9).value = code

                rep = auburn.cell(row=i, column=12).value
                rep = rep.strip().lower()
                monthly.cell(row=num, column=6).value = rep

                # setting commision for pete
                if rep == "maie":
                    monthly.cell(row=num, column=7).value = None
                elif rep == "pete code lead" and pete_commission() > 5:
                    monthly.cell(row=num, column=7).value = 75
                elif rep == "pete code lead" and pete_commission() <= 5:
                    monthly.cell(row=num, column=7).value = 'x'
                else:
                    monthly.cell(
                        row=num, column=7).value = set_commission(course)

                vender = auburn.cell(row=i, column=13).value
                jon_sheet.cell(row=num1, column=4).value = vender
                if vender == 'CCI':
                    if course in au_programs:
                        monthly.cell(row=num, column=9).value = 'AU'
                        monthly.cell(
                            row=num, column=set_pricing_column(
                                'AU')).value = set_pricing_au(course)
                    elif course not in au_programs:
                        monthly.cell(row=num, column=9).value = 'MET'
                        monthly.cell(row=num, column=set_pricing_column(
                            'MET')).value = set_pricing_met(course)
                elif vender == 'Pete Medd':
                    monthly.cell(row=num, column=9).value = 'AU M'
                    monthly.cell(row=num, column=set_pricing_column(
                        'MET')).value = set_pricing_met(course)
                else:
                    monthly.cell(row=num, column=9).value = 'AU ED4'
                    monthly.cell(row=num, column=set_pricing_column(
                        'AU ED4')).value = set_pricing_met(course)

                num += 1
                num1 += 1


# -----------------------------------------------------------------------------Clemson-----------------


def clemson_students(current_month):
    mr = clemson.max_row
    mc = clemson.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(450, mr+1):

        c = clemson.cell(row=i, column=3).value
        name = clemson.cell(row=i, column=9).value
        last_number_row = num - 1
        if c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:
                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = clemson.cell(row=i, column=1).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1

                date2 = clemson.cell(row=i, column=3).value
                date2 = date2.strftime('%m') + '/' + \
                    date2.strftime('%d') + '/' + date2.strftime('%-y')
                monthly.cell(row=num, column=3).value = date2

                date3 = clemson.cell(row=i, column=4).value
                date3 = date3.strftime('%m') + '/' + \
                    date3.strftime('%d') + '/' + date3.strftime('%-y')
                jon_sheet.cell(row=num1, column=3).value = date3

                address = clemson.cell(row=i, column=7).value
                monthly.cell(row=num, column=14).value = address

                email = clemson.cell(row=i, column=8).value
                jon_sheet.cell(row=num1, column=2).value = email

                if 'LAPTOP' in name:
                    monthly.cell(row=num, column=8).value = 'x'
                monthly.cell(row=num, column=4).value = name
                jon_sheet.cell(row=num1, column=1).value = name

                course = clemson.cell(row=i, column=10).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                code = clemson.cell(row=i, column=11).value
                monthly.cell(row=num, column=9).value = code

                rep = clemson.cell(row=i, column=12).value
                rep = rep.strip().lower()
                monthly.cell(row=num, column=6).value = rep

                # setting commision for pete
                if rep == "maie":
                    monthly.cell(row=num, column=7).value = None
                elif rep == "pete code lead" and pete_commission() > 5:
                    monthly.cell(row=num, column=7).value = 75
                elif rep == "pete code lead" and pete_commission() <= 5:
                    monthly.cell(row=num, column=7).value = 'x'
                else:
                    monthly.cell(
                        row=num, column=7).value = set_commission(course)

                vender = clemson.cell(row=i, column=13).value
                jon_sheet.cell(row=num1, column=4).value = vender
                monthly.cell(row=num, column=9).value = 'CLEM'
                monthly.cell(row=num, column=set_pricing_column(
                    'CLEM')).value = set_pricing_cci(course)

                num += 1
                num1 += 1

# -----------------------------------------------------------------------------CSU-----------------


def csu_students(current_month):
    mr = csu.max_row
    mc = csu.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(96, mr+1):

        c = csu.cell(row=i, column=3).value
        name = csu.cell(row=i, column=9).value
        last_number_row = num - 1
        if c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:
                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = csu.cell(row=i, column=1).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1

                date2 = csu.cell(row=i, column=3).value
                date2 = date2.strftime('%m') + '/' + \
                    date2.strftime('%d') + '/' + date2.strftime('%-y')
                monthly.cell(row=num, column=3).value = date2

                date3 = csu.cell(row=i, column=4).value
                date3 = date3.strftime('%m') + '/' + \
                    date3.strftime('%d') + '/' + date3.strftime('%-y')
                jon_sheet.cell(row=num1, column=3).value = date3

                address = csu.cell(row=i, column=7).value
                monthly.cell(row=num, column=14).value = address

                email = csu.cell(row=i, column=8).value
                jon_sheet.cell(row=num1, column=2).value = email

                if 'LAPTOP' in name:
                    monthly.cell(row=num, column=8).value = 'x'
                monthly.cell(row=num, column=4).value = name
                jon_sheet.cell(row=num1, column=1).value = name

                course = csu.cell(row=i, column=10).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                code = csu.cell(row=i, column=11).value
                monthly.cell(row=num, column=9).value = code

                rep = csu.cell(row=i, column=12).value
                rep = rep.strip().lower()
                monthly.cell(row=num, column=6).value = rep

                # setting commision for pete
                if rep == "maie":
                    monthly.cell(row=num, column=7).value = None
                elif rep == "pete code lead" and pete_commission() > 5:
                    monthly.cell(row=num, column=7).value = 75
                elif rep == "pete code lead" and pete_commission() <= 5:
                    monthly.cell(row=num, column=7).value = 'x'
                else:
                    monthly.cell(
                        row=num, column=7).value = set_commission(course)

                vender = csu.cell(row=i, column=13).value
                jon_sheet.cell(row=num1, column=4).value = vender
                monthly.cell(row=num, column=9).value = 'CSU'
                monthly.cell(row=num, column=set_pricing_column(
                    'CSU')).value = set_pricing_csu(course)

                num += 1
                num1 += 1

# -----------------------------------------------------------------------------LSU-----------------


def lsu_students(current_month):
    mr = lsu.max_row
    mc = lsu.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(74, mr+1):

        c = lsu.cell(row=i, column=3).value
        name = lsu.cell(row=i, column=9).value
        last_number_row = num - 1
        if c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:
                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = lsu.cell(row=i, column=1).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1

                date2 = lsu.cell(row=i, column=3).value
                date2 = date2.strftime('%m') + '/' + \
                    date2.strftime('%d') + '/' + date2.strftime('%-y')
                monthly.cell(row=num, column=3).value = date2

                date3 = lsu.cell(row=i, column=4).value
                date3 = date3.strftime('%m') + '/' + \
                    date3.strftime('%d') + '/' + date3.strftime('%-y')
                jon_sheet.cell(row=num1, column=3).value = date3

                address = lsu.cell(row=i, column=7).value
                monthly.cell(row=num, column=14).value = address

                email = lsu.cell(row=i, column=8).value
                jon_sheet.cell(row=num1, column=2).value = email

                if 'LAPTOP' in name:
                    monthly.cell(row=num, column=8).value = 'x'
                monthly.cell(row=num, column=4).value = name
                jon_sheet.cell(row=num1, column=1).value = name

                course = lsu.cell(row=i, column=10).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                code = lsu.cell(row=i, column=11).value
                monthly.cell(row=num, column=9).value = code

                rep = lsu.cell(row=i, column=12).value
                rep = rep.strip().lower()
                monthly.cell(row=num, column=6).value = rep

                # setting commision for pete
                if rep == "maie":
                    monthly.cell(row=num, column=7).value = None
                elif rep == "pete code lead" and pete_commission() > 5:
                    monthly.cell(row=num, column=7).value = 75
                elif rep == "pete code lead" and pete_commission() <= 5:
                    monthly.cell(row=num, column=7).value = 'x'
                else:
                    monthly.cell(
                        row=num, column=7).value = set_commission(course)

                vender = lsu.cell(row=i, column=13).value
                jon_sheet.cell(row=num1, column=4).value = vender
                monthly.cell(row=num, column=9).value = 'LSU'
                monthly.cell(row=num, column=set_pricing_column(
                    'LSU')).value = set_pricing_cci(course)

                num += 1
                num1 += 1
# -----------------------------------------------------------------------------MSU-----------------


def msu_students(current_month):
    mr = msu.max_row
    mc = msu.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(450, mr+1):

        c = msu.cell(row=i, column=3).value
        name = msu.cell(row=i, column=9).value
        last_number_row = num - 1
        if c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:
                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = msu.cell(row=i, column=1).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1

                date2 = msu.cell(row=i, column=3).value
                date2 = date2.strftime('%m') + '/' + \
                    date2.strftime('%d') + '/' + date2.strftime('%-y')
                monthly.cell(row=num, column=3).value = date2

                date3 = msu.cell(row=i, column=4).value
                date3 = date3.strftime('%m') + '/' + \
                    date3.strftime('%d') + '/' + date3.strftime('%-y')
                jon_sheet.cell(row=num1, column=3).value = date3

                address = msu.cell(row=i, column=7).value
                monthly.cell(row=num, column=14).value = address

                email = msu.cell(row=i, column=8).value
                jon_sheet.cell(row=num1, column=2).value = email

                if 'LAPTOP' in name:
                    monthly.cell(row=num, column=8).value = 'x'
                monthly.cell(row=num, column=4).value = name
                jon_sheet.cell(row=num1, column=1).value = name

                course = msu.cell(row=i, column=10).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                code = msu.cell(row=i, column=11).value
                monthly.cell(row=num, column=9).value = code

                rep = msu.cell(row=i, column=12).value
                rep = rep.strip().lower()
                monthly.cell(row=num, column=6).value = rep

                # setting commision for pete
                if rep == "maie":
                    monthly.cell(row=num, column=7).value = None
                elif rep == "pete code lead" and pete_commission() > 5:
                    monthly.cell(row=num, column=7).value = 75
                elif rep == "pete code lead" and pete_commission() <= 5:
                    monthly.cell(row=num, column=7).value = 'x'
                else:
                    monthly.cell(
                        row=num, column=7).value = set_commission(course)

                vender = msu.cell(row=i, column=13).value
                jon_sheet.cell(row=num1, column=4).value = vender
                monthly.cell(row=num, column=9).value = 'MSU'
                monthly.cell(row=num, column=set_pricing_column(
                    'MSU')).value = set_pricing_cci(course)

                num += 1
                num1 += 1

# -----------------------------------------------------------------------------UNH-----------------


def unh_students(current_month):
    mr = unh.max_row
    mc = unh.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(26, mr+1):

        c = unh.cell(row=i, column=3).value
        name = unh.cell(row=i, column=9).value
        last_number_row = num - 1
        if c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:
                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = unh.cell(row=i, column=1).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1

                date2 = unh.cell(row=i, column=3).value
                date2 = date2.strftime('%m') + '/' + \
                    date2.strftime('%d') + '/' + date2.strftime('%-y')
                monthly.cell(row=num, column=3).value = date2

                date3 = unh.cell(row=i, column=4).value
                date3 = date3.strftime('%m') + '/' + \
                    date3.strftime('%d') + '/' + date3.strftime('%-y')
                jon_sheet.cell(row=num1, column=3).value = date3

                address = unh.cell(row=i, column=7).value
                monthly.cell(row=num, column=14).value = address

                email = unh.cell(row=i, column=8).value
                jon_sheet.cell(row=num1, column=2).value = email

                if 'LAPTOP' in name:
                    monthly.cell(row=num, column=8).value = 'x'
                monthly.cell(row=num, column=4).value = name
                jon_sheet.cell(row=num1, column=1).value = name

                course = unh.cell(row=i, column=10).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                code = unh.cell(row=i, column=11).value
                monthly.cell(row=num, column=9).value = code

                rep = unh.cell(row=i, column=12).value
                rep = rep.strip().lower()
                monthly.cell(row=num, column=6).value = rep

                # setting commision for pete
                if rep == "maie":
                    monthly.cell(row=num, column=7).value = None
                elif rep == "pete code lead" and pete_commission() > 5:
                    monthly.cell(row=num, column=7).value = 75
                elif rep == "pete code lead" and pete_commission() <= 5:
                    monthly.cell(row=num, column=7).value = 'x'
                else:
                    monthly.cell(
                        row=num, column=7).value = set_commission(course)

                vender = unh.cell(row=i, column=13).value
                jon_sheet.cell(row=num1, column=4).value = vender
                monthly.cell(row=num, column=9).value = 'UNH'
                monthly.cell(row=num, column=set_pricing_column(
                    'UNH')).value = set_pricing_cci(course)

                num += 1
                num1 += 1
# -----------------------------------------------------------------------------TAMUT-----------------


def tamu_students(current_month):
    mr = tamu.max_row
    mc = tamu.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(50, mr+1):

        c = tamu.cell(row=i, column=3).value
        name = tamu.cell(row=i, column=9).value
        last_number_row = num - 1
        if c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:
                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = tamu.cell(row=i, column=1).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1

                date2 = tamu.cell(row=i, column=3).value
                date2 = date2.strftime('%m') + '/' + \
                    date2.strftime('%d') + '/' + date2.strftime('%-y')
                monthly.cell(row=num, column=3).value = date2

                date3 = tamu.cell(row=i, column=4).value
                date3 = date3.strftime('%m') + '/' + \
                    date3.strftime('%d') + '/' + date3.strftime('%-y')
                jon_sheet.cell(row=num1, column=3).value = date3

                address = tamu.cell(row=i, column=7).value
                monthly.cell(row=num, column=14).value = address

                email = tamu.cell(row=i, column=8).value
                jon_sheet.cell(row=num1, column=2).value = email

                if 'LAPTOP' in name:
                    monthly.cell(row=num, column=8).value = 'x'
                monthly.cell(row=num, column=4).value = name
                jon_sheet.cell(row=num1, column=1).value = name

                course = tamu.cell(row=i, column=10).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                code = tamu.cell(row=i, column=11).value
                monthly.cell(row=num, column=9).value = code

                rep = tamu.cell(row=i, column=12).value
                rep = rep.strip().lower()
                monthly.cell(row=num, column=6).value = rep

                # setting commision for pete
                if rep == "maie":
                    monthly.cell(row=num, column=7).value = None
                elif rep == "pete code lead" and pete_commission() > 5:
                    monthly.cell(row=num, column=7).value = 75
                elif rep == "pete code lead" and pete_commission() <= 5:
                    monthly.cell(row=num, column=7).value = 'x'
                else:
                    monthly.cell(
                        row=num, column=7).value = set_commission(course)

                vender = tamu.cell(row=i, column=13).value
                jon_sheet.cell(row=num1, column=4).value = vender
                monthly.cell(row=num, column=9).value = 'TAMU'
                monthly.cell(row=num, column=set_pricing_column(
                    'TAMU')).value = set_pricing_cci(course)

                num += 1
                num1 += 1
# -----------------------------------------------------------------------------WKU-----------------


def wku_students(current_month):
    mr = wku.max_row
    mc = wku.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(257, mr+1):

        c = wku.cell(row=i, column=3).value
        name = wku.cell(row=i, column=9).value
        last_number_row = num - 1
        if c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:
                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = wku.cell(row=i, column=1).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1

                date2 = wku.cell(row=i, column=3).value
                date2 = date2.strftime('%m') + '/' + \
                    date2.strftime('%d') + '/' + date2.strftime('%-y')
                monthly.cell(row=num, column=3).value = date2

                date3 = wku.cell(row=i, column=4).value
                date3 = date3.strftime('%m') + '/' + \
                    date3.strftime('%d') + '/' + date3.strftime('%-y')
                jon_sheet.cell(row=num1, column=3).value = date3

                address = wku.cell(row=i, column=7).value
                monthly.cell(row=num, column=14).value = address

                email = wku.cell(row=i, column=8).value
                jon_sheet.cell(row=num1, column=2).value = email

                if 'LAPTOP' in name:
                    monthly.cell(row=num, column=8).value = 'x'
                monthly.cell(row=num, column=4).value = name
                jon_sheet.cell(row=num1, column=1).value = name

                course = wku.cell(row=i, column=10).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                code = wku.cell(row=i, column=11).value
                monthly.cell(row=num, column=9).value = code

                rep = wku.cell(row=i, column=12).value
                rep = rep.strip().lower()
                monthly.cell(row=num, column=6).value = rep

                # setting commision for pete
                if rep == "maie":
                    monthly.cell(row=num, column=7).value = None
                elif rep == "pete code lead" and pete_commission() > 5:
                    monthly.cell(row=num, column=7).value = 75
                elif rep == "pete code lead" and pete_commission() <= 5:
                    monthly.cell(row=num, column=7).value = 'x'
                else:
                    monthly.cell(
                        row=num, column=7).value = set_commission(course)

                vender = wku.cell(row=i, column=13).value
                jon_sheet.cell(row=num1, column=4).value = vender
                monthly.cell(row=num, column=9).value = 'WKU'
                monthly.cell(row=num, column=set_pricing_column(
                    'WKU')).value = set_pricing_cci(course)

                num += 1
                num1 += 1
# -----------------------------------------------------------------------------UWLAX-----------------


def uwlax_students(current_month):
    mr = uwlax.max_row
    mc = uwlax.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(246, mr+1):

        c = uwlax.cell(row=i, column=3).value
        name = uwlax.cell(row=i, column=9).value
        last_number_row = num - 1
        if c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:
                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = uwlax.cell(row=i, column=1).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1

                date2 = uwlax.cell(row=i, column=3).value
                date2 = date2.strftime('%m') + '/' + \
                    date2.strftime('%d') + '/' + date2.strftime('%-y')
                monthly.cell(row=num, column=3).value = date2

                date3 = uwlax.cell(row=i, column=4).value
                date3 = date3.strftime('%m') + '/' + \
                    date3.strftime('%d') + '/' + date3.strftime('%-y')
                jon_sheet.cell(row=num1, column=3).value = date3

                address = uwlax.cell(row=i, column=7).value
                monthly.cell(row=num, column=14).value = address

                email = uwlax.cell(row=i, column=8).value
                jon_sheet.cell(row=num1, column=2).value = email

                if 'LAPTOP' in name:
                    monthly.cell(row=num, column=8).value = 'x'
                monthly.cell(row=num, column=4).value = name
                jon_sheet.cell(row=num1, column=1).value = name

                course = uwlax.cell(row=i, column=10).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                code = uwlax.cell(row=i, column=11).value
                monthly.cell(row=num, column=9).value = code

                rep = uwlax.cell(row=i, column=12).value
                rep = rep.strip().lower()
                monthly.cell(row=num, column=6).value = rep

                # setting commision for pete
                if rep == "maie":
                    monthly.cell(row=num, column=7).value = None
                elif rep == "pete code lead" and pete_commission() > 5:
                    monthly.cell(row=num, column=7).value = 75
                elif rep == "pete code lead" and pete_commission() <= 5:
                    monthly.cell(row=num, column=7).value = 'x'
                else:
                    monthly.cell(
                        row=num, column=7).value = set_commission(course)

                vender = uwlax.cell(row=i, column=13).value
                jon_sheet.cell(row=num1, column=4).value = vender
                monthly.cell(row=num, column=9).value = 'UWLAX'
                monthly.cell(row=num, column=set_pricing_column(
                    'UWLAX')).value = set_pricing_uwlax(course)

                num += 1
                num1 += 1
# -----------------------------------------------------------------------------DESU-----------------


def desu_students(current_month):
    mr = desu.max_row
    mc = desu.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(11, mr+1):

        c = desu.cell(row=i, column=3).value
        name = desu.cell(row=i, column=9).value
        last_number_row = num - 1
        if c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:
                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = desu.cell(row=i, column=1).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1

                date2 = desu.cell(row=i, column=3).value
                date2 = date2.strftime('%m') + '/' + \
                    date2.strftime('%d') + '/' + date2.strftime('%-y')
                monthly.cell(row=num, column=3).value = date2

                date3 = desu.cell(row=i, column=4).value
                date3 = date3.strftime('%m') + '/' + \
                    date3.strftime('%d') + '/' + date3.strftime('%-y')
                jon_sheet.cell(row=num1, column=3).value = date3

                address = desu.cell(row=i, column=7).value
                monthly.cell(row=num, column=14).value = address

                email = desu.cell(row=i, column=8).value
                jon_sheet.cell(row=num1, column=2).value = email

                if 'LAPTOP' in name:
                    monthly.cell(row=num, column=8).value = 'x'
                monthly.cell(row=num, column=4).value = name
                jon_sheet.cell(row=num1, column=1).value = name

                course = desu.cell(row=i, column=10).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                code = desu.cell(row=i, column=11).value
                monthly.cell(row=num, column=9).value = code

                rep = desu.cell(row=i, column=12).value
                rep = rep.strip().lower()
                monthly.cell(row=num, column=6).value = rep

                # setting commision for pete
                if rep == "maie":
                    monthly.cell(row=num, column=7).value = None
                elif rep == "pete code lead" and pete_commission() > 5:
                    monthly.cell(row=num, column=7).value = 75
                elif rep == "pete code lead" and pete_commission() <= 5:
                    monthly.cell(row=num, column=7).value = 'x'
                else:
                    monthly.cell(
                        row=num, column=7).value = set_commission(course)

                vender = desu.cell(row=i, column=13).value
                jon_sheet.cell(row=num1, column=4).value = vender
                monthly.cell(row=num, column=9).value = 'DESU'
                monthly.cell(row=num, column=set_pricing_column(
                    'DESU')).value = set_pricing_cci(course)

                num += 1
                num1 += 1


# Dictionaries
cci_programs = dict({
    "accounting professional": 2016,
    "administrative assistant with quickbooks": 1748,
    "bookkeeping with quickbooks": 1733,
    "childcare specialist": 1942,
    "clinical medical assistant": 1405,
    "clinical medical assistant with ob/gyn": 1405,
    "clinical medical assistant with pediatric specialist": 1405,
    "criminal investigation professional": 2051,
    "dental assisting": 1825,
    "ekg technician cert program": 1250,
    "finanance professional": 1967,
    "human resources professional": 2029,
    "it cyber security professional with comp tia security +": 2050,
    "medical administration assistance": 1250,
    "medical billing and coding": 1215,
    "medical billing and coding with medical administrative assistant": 1370,
    "medical billing and coding with medical admin": 1370,
    "medical billing and coding with medical administration": 1370,
    "medical billing & coding w/ medical administrative assistant certificate program includes cmaa and cpc national certification exams": 1370,
    "paralegal": 1699,
    "pharmacy technician": 1200,
    "pharmacy technician with medical administration": 1400,
    "phlebotomy technician": 1575,
    "photography entrepreneur with adobe certificate": 1850,
    "teachers aide": 2029,
    "veterinary assistant specialist": 1013, })

au_programs = dict({
    "clinical medical assistant": 1405,
    "clinical medical assistant with ob/gyn": 1405,
    "clinical medical assistant with pediatric specialist": 1405,
    "dental assisting": 1825,
    "ekg technician cert program": 1250,
    "medical billing and coding": 1215,
    "medical billing and coding with medical administrative assistant": 1370,
    "medical billing and coding with medical admin": 1370,
    "medical billing and coding with medical administration": 1370,
    "medical billing & coding w/ medical administrative assistant certificate program includes cmaa and cpc national certification exams": 1370,
    "pharmacy technician": 1200,
    "pharmacy technician with medical administration": 1400,
    "phlebotomy technician": 1575,
    "veterinary assistant specialist": 1013, })

met_programs = dict({
    "accounting professional": 2999.25,
    "administrative assistant with quickbooks": 2999.25,
    "bookeeping with quickbooks": 2849.25,
    "business management professional": 2999.25,
    "childcare specialist": 2999.25,
    "child day care management cert program": 2962.50,
    "event planning entrepreneur": 2962.50,
    "full stack web developer with mean stack": 2999.25,
    "human resources professional": 2999.25,
    "life skills coach": 2962.50,
    "massage practitioner program (500 hr)": 4874.25,
    "massage practitioner program (620 hr)": 5249.25,
    "massage practitioner program (650 hr)": 5474.25,
    "massage practitioner program (700 hr)": 6524.25,
    "massage practitioner program (750 hr)": 5999.25,
    "marketing professional": 2849.25,
    "mental health technician specialist cert": 2962.50,
    "ophthalmic assistant specialist": 2962.50,
    "paralegal certificate program": 2999.25,
    "patient advocate specialist": 2962.50,
    "personal fitness trainer specialist": 2999.25,
    "photography entrepreneur with adobe certificate": 2962.50,
    "photography entrepreneur with adobe": 2962.50,
    "physical therapy aide": 2962.50,
    "teachers aide": 2999.25})

uwlax_programs = dict({
    "clinical medical assistant": 2765,
    "dental assisting": 2765,
    "dental assisting certification": 2765,
    "medical billing and coding with medical admin": 2765,
    "teachers aide": 2799.30})

csu_programs = dict({
    "clinical medical assistant": 2962.50,
    "medical billing and coding": 2437.50,
    "medical billing and coding with medical administration": 2962.50})

commission = dict({
    "accounting professional": 500,
    "administrative assistant with quickbooks": 500,
    "bookkeeping with quickbooks": 400,
    "child day care management cert program": 400,
    "childcare specialist": 500,
    "clinical medical assistant": 300,
    "clinical medical assistant with ob/gyn": 300,
    "clinical medical assistant with pediatric specialist": 300,
    "criminal investigation professional": 500,
    "dental assisting certification": 400,
    "dental assisting": 400,
    "life skills coach": 300,
    "ekg technician cert program": 300,
    "event planning entrepreneur": 300,
    "finanance professional": 500,
    "full stack web developer with mean stack": None,
    "human resources professional": 500,
    "it cyber security professional with comp tia security +": 500,
    "medical administration assistance": None,
    "medical billing and coding": 300,
    "medical billing and coding with medical administration": 300,
    "medical billing and coding with medical admin": 300,
    "medical billing and coding with medical administration": 300,
    "medical billing & coding w/ medical administrative assistant certificate program includes cmaa and cpc national certification exams": 300,
    "mental health technician specialist cert": 400,
    "paralegal": 500,
    "pharmacy technician": 300,
    "pharmacy technician with medical administration": 300,
    "phlebotomy technician": 500,
    "photography entrepreneur with adobe certificate": 300,
    "photography entrepreneur with adobe": 300,
    "patient advocate specialist": 400,
    "teachers aide": 500,
    "veterinary assistant specialist": 300
})


def set_pricing_cci(course):
    if course in cci_programs:

        return cci_programs[course]
    else:
        print("No class, update pricing for cci")


def set_pricing_au(course):
    if course in au_programs:

        return au_programs[course]
    else:
        print("No class update pricing for au")


def set_pricing_met(course):
    if course in met_programs:
        return met_programs[course]
    else:
        print("No class update pricing for met")


def set_pricing_uwlax(course):
    if course in uwlax_programs:
        return uwlax_programs[course]
    else:
        print("No class update pricing for uwlax")


def set_pricing_csu(course):
    if course in csu_programs:
        return csu_programs[course]
    else:
        print("No class, update pricing for csu")


def set_pricing_column(school):

    if school == "AU ED4":
        return 12
    elif school == "AU":
        return 13
    elif school == "MET":
        return 12
    elif school == "WKU":
        return 17
    elif school == "LSU":
        return 19
    elif school == "CLEM":
        return 21
    elif school == "UWLAX":
        return 22
    elif school == "CSU":
        return 25
    elif school == "TAMU":
        return 26
    elif school == "MSU":
        return 28
    elif school == "UNH":
        return 31
    elif school == "DESU":
        return 32
    else:
        print("\033[1;32mno school with that name \033[0;0m")


def pete_commission():
    num = 0
    for cell in monthly["F"]:
        if cell.value == "pete code lead":
            num += 1
    return num


def set_commission(course):
    if course in commission:
        return commission[course]
    else:
        print('\033[1;32mno class to set commission \033[0;0m')


def runProgram():
    start = findNextCell()
    auburn_students('04')
    clemson_students('04')
    csu_students('04')
    lsu_students('04')
    msu_students('04')
    unh_students('04')
    tamu_students('04')
    wku_students('04')
    uwlax_students('04')
    desu_students('04')
    wb2.save(monthly_spreadsheet)
    wb3.save(jon_email_workbook)
    end = findNextCell()
    total = end-start
    print("All Done Transferring Students!")
    print("\033[1;32m{} \033[0;0mwere transferred".format(total))

