import openpyxl as xl
from datetime import datetime
import os
from difflib import SequenceMatcher
import Levenshtein

jon_email_workbook = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Letters to students/Weekly Email for Lisa/Jon weekly email list.xlsx"
pete_spreadsheet = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Pete's Backup/MILTARY/PETE ALL 3 SPREADSHEETS MYCAA FOR STACEY AND LISA/MAIN ENROLLMENT FOLDER/SPREADSHEETS/students mycaa FINAL-TODAY.xlsx"
monthly_spreadsheet = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Invoices/2020 Enrollment/June 2020.xlsx"
# assert os.path.exists(pete_spreadsheet)

lastMonth = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Invoices/2020 Enrollment/May 2020.xlsx"
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
tamiu = wb1.worksheets[12]

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


def findMissingClass(dictionary, wrong):
    num = 0
    name = ''
    for key in dictionary:
        # print(SequenceMatcher(None, key, 'veterinary assistant specialist').ratio())
        ratio = Levenshtein.ratio(key, wrong)
        if ratio > num:
            num = ratio
            name = key
    if num > 0.5:
        print(
            f'Smart lookup finished. \033[1;32m{num}%\033[0;0m that \033[1;33m{wrong}\033[0;0m is \033[1;32m{name}\033[0;0m')
        return dictionary[name]
    else:
        print("Smart lookup finished. Nothing really seems to match")
        return dictionary[name]


def auburn_students(current_month):
    mr = auburn.max_row
    mc = auburn.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(6060, mr+1):

        c = auburn.cell(row=i, column=3).value

        name = auburn.cell(row=i, column=9).value

        last_number_row = num - 1

        if c != None and c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
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
                res = isinstance(date3, datetime)

                if res:
                    date3 = date3.strftime('%m') + '/' + \
                        date3.strftime('%d') + '/' + date3.strftime('%-y')
                    jon_sheet.cell(row=num1, column=3).value = date3
                else:
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

    wb2.save(monthly_spreadsheet)


# -----------------------------------------------------------------------------Other Schools-----------------


def school_tab(current_month, school, schoolString, rowNumber):
    mr = school.max_row
    mc = school.max_column
    num = findNextCell()
    num1 = findNextCellJonEmail()

    for i in range(rowNumber, mr+1):

        c = school.cell(row=i, column=3).value
        name = school.cell(row=i, column=9).value
        last_number_row = num - 1
        if c != None and c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
            if findName(name) != True:
                # place invoice number
                last_invoice_number = monthly.cell(
                    row=last_number_row, column=11).value
                monthly.cell(row=num, column=11).value = last_invoice_number+1
                # place first date
                date1 = school.cell(row=i, column=1).value
                date1 = date1.strftime('%m') + '/' + \
                    date1.strftime('%d') + '/' + date1.strftime('%-y')
                monthly.cell(row=num, column=2).value = date1

                date2 = school.cell(row=i, column=3).value
                date2 = date2.strftime('%m') + '/' + \
                    date2.strftime('%d') + '/' + date2.strftime('%-y')
                monthly.cell(row=num, column=3).value = date2

                date3 = school.cell(row=i, column=4).value
                date3 = date3.strftime('%m') + '/' + \
                    date3.strftime('%d') + '/' + date3.strftime('%-y')
                jon_sheet.cell(row=num1, column=3).value = date3

                address = school.cell(row=i, column=7).value
                monthly.cell(row=num, column=14).value = address

                email = school.cell(row=i, column=8).value
                jon_sheet.cell(row=num1, column=2).value = email

                if 'LAPTOP' in name:
                    monthly.cell(row=num, column=8).value = 'x'
                monthly.cell(row=num, column=4).value = name
                jon_sheet.cell(row=num1, column=1).value = name

                course = school.cell(row=i, column=10).value
                course = course.strip().lower()
                monthly.cell(row=num, column=5).value = course

                code = school.cell(row=i, column=11).value
                monthly.cell(row=num, column=9).value = code

                rep = school.cell(row=i, column=12).value
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

                vender = school.cell(row=i, column=13).value
                jon_sheet.cell(row=num1, column=4).value = vender
                monthly.cell(row=num, column=9).value = schoolString
                if schoolString == 'CSU':
                    monthly.cell(row=num, column=set_pricing_column(
                        schoolString)).value = set_pricing_csu(course)
                elif schoolString == 'UWLAX':
                    monthly.cell(row=num, column=set_pricing_column(
                        schoolString)).value = set_pricing_uwlax(course)
                else:
                    monthly.cell(row=num, column=set_pricing_column(
                        schoolString)).value = set_pricing_cci(course)

                num += 1
                num1 += 1

    wb2.save(monthly_spreadsheet)


# Dictionaries
cci_programs = dict({
    "accounting professional": 2016,
    "administrative assistant with quickbooks": 1748,
    "bookkeeping with quickbooks": 1733,
    "business management professional": 1948,
    "childcare specialist": 1942,
    "clinical medical assistant": 1405,
    "clinical medical assistant with ob/gyn": 1405,
    "clinical medical assistant with pediatric specialist": 1405,
    "criminal investigation professional": 2051,
    "dental assisting": 1825,
    "ekg technician cert program": 1250,
    "finanance professional": 1967,
    "finance professional": 1967,
    "front end web developer": 1850,
    "human resources professional": 2029,
    "it cyber security professional with comp tia security +": 2050,
    "medical administration assistance": 1250,
    "medical administrative assistant": 1250,
    "medical administrative assistant online": 1250,
    "medical administrative assistant online": 1250,
    "medical billing and coding": 1215,
    "medical billing and coding with medical administrative assistant": 1370,
    "medical billing and coding with medical admin": 1370,
    "medical billing and coding with medical administration": 1370,
    "medical billing & coding w/ medical administrative assistant certificate program includes cmaa and cpc national certification exams": 1370,
    "organizational behavior professional": 2090,
    "paralegal": 1699,
    "paralegal certificate program": 1699,
    "pharmacy technician": 1200,
    "pharmacy technician with medical administration": 1400,
    "pharmacy tech with med admin": 1400,
    "phlebotomy technician": 1575,
    "phlebotomy tech -spanish": 1575,
    "photography entrepreneur with adobe certificate": 1850,
    "photography entrepreneur with adobe": 1850,
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
    "pharmacy technician with medical admin online inc. national cert. and clinical ext.": 1400,
    "phlebotomy technician": 1575,
    "veterinary assistant specialist": 1013, })

met_programs = dict({
    "accounting professional": 2999.25,
    "administrative assistant with quickbooks": 2999.25,
    "administrative assistant with bookkeeping and quickbooks": 2999.25,
    "bookeeping with quickbooks": 2849.25,
    "business management professional": 2999.25,
    "childcare specialist": 2999.25,
    "child day care management cert program": 2962.50,
    "drug and alcohol counselor": 2962.50,
    "event planning entrepreneur": 2962.50,
    "full stack web developer with mean stack": 2999.25,
    "front end web developer": 2999.25,
    "health & fitness industry professional": 2962.50,
    "homeland security specialist": 2849.25,
    "human resources professional": 2999.25,
    "interior decorating and design entrepreneur": 2962.50,
    "it cyber security professional with comp tia security+": 2999.25,
    "it network professional with comptia network+": 2999.25,
    "life skills coach": 2962.50,
    "massage practitioner program (500 hr)": 3000,
    "massage practitioner program (620 hr)": 3000,
    "massage practitioner program (650 hr)": 3000,
    "massage practitioner program (700 hr)": 3000,
    "massage practitioner program (750 hr)": 3000,
    "marketing professional": 2849.25,
    "mental health technician specialist cert": 2962.50,
    "nutrition and fitness professional": 2962.50,
    "ophthalmic assistant specialist": 2962.50,
    "paralegal certificate program": 2999.25,
    "patient advocate specialist": 2962.50,
    "personal fitness trainer specialist": 2999.25,
    "photography entrepreneur with adobe certificate": 2962.50,
    "photography entrepreneur with adobe": 2962.50,
    "physical therapy aide": 2962.50,
    "professional cooking and catering": 2962.50,
    "real estate law professional": 2849.25,
    "stress management coach": 2962.50,
    "teachers aide": 2999.25,
    "technical writing": 1649.25,
    "travel agent specialist": 2962.50,
    'veterinary office assistant specialist': 2962.50,
    "wedding consultant entrepreneur": 2962.50})

uwlax_programs = dict({
    "clinical medical assistant": 2765,
    "dental assisting": 2765,
    "dental assisting certification": 2765,
    "medical administrative assistant": 2100,
    "medical billing and coding with medical admin": 2765,
    "pharmacy technician with medical administration": 2765,
    "physicians office assistant with ehrm": 2765,
    "teachers aide": 2799.30,
    "veterinary assistant": 2695})

csu_programs = dict({
    "clinical medical assistant": 2962.50,
    "medical billing and coding": 2437.50,
    "medical billing and coding with medical administration": 2962.50,
    "paralegal": 2999.25})

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
        print(cci_programs[course])
    else:
        print("\033[1;31mNo class found for CCI, smart update started... \033[0;0m")
        return findMissingClass(cci_programs, course)


def set_pricing_au(course):
    if course in au_programs:

        return au_programs[course]
    else:
        print(
            "\033[1;31mNo class update pricing for au, smart update started... \033[0;0m")
        return findMissingClass(au_programs, course)


def set_pricing_met(course):
    if course in met_programs:
        return met_programs[course]
    else:
        print(
            "\033[1;31mNo class update pricing for met, smart update started... \033[0;0m")
        return findMissingClass(met_programs, course)


def set_pricing_uwlax(course):
    if course in uwlax_programs:
        return uwlax_programs[course]
    else:
        print(
            "\033[1;31mNo class update pricing for uwlax, smart update started... \033[0;0m")
        return findMissingClass(uwlax_programs, course)


def set_pricing_csu(course):
    if course in csu_programs:
        return csu_programs[course]
    else:
        print(
            "\033[1;31mNo class, update pricing for csu, smart update started... \033[0;0m")
        return findMissingClass(csu_programs, course)


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
    elif school == "TAMIU":
        return 33
    else:
        print("\033[1;31mno school with that name \033[0;0m")


def pete_commission():
    num = 0
    for cell in monthly["F"]:
        if cell.value == "pete code lead" or 'pete':
            num += 1
    return num


def set_commission(course):
    if course in commission:
        return commission[course]
    else:
        print('')


def runProgram():
    start = findNextCell()
    auburn_students('06')
    school_tab('06', clemson, 'CLEM', 450)
    school_tab('06', csu, 'CSU', 96)
    school_tab('06', lsu, 'LSU', 74)
    school_tab('06', msu, 'MSU', 450)
    school_tab('06', unh, 'UNH', 26)
    school_tab('06', tamu, 'TAMU', 50)
    school_tab('06', wku, 'WKU', 257)
    school_tab('06', uwlax, 'UWLAX', 246)
    school_tab('06', desu, 'DESU', 11)
    school_tab('06', tamiu, 'TAMIU', 9)
    wb2.save(monthly_spreadsheet)
    wb3.save(jon_email_workbook)
    end = findNextCell()
    total = end-start
    print("\033[1;32mAll Done Transferring Students!\033[0;0m")
    print("\033[1;32m{} \033[0;0mwere transferred".format(total))


# findNextCellPete(auburn)
# set_pricing_cci('Veteriary Assistant')
