import json
import boto3
import botocore
import dropbox
import openpyxl as xl
import os
from datetime import datetime
import Levenshtein
import time
from data import cci_programs, commission, au_programs, met_programs, uwlax_programs, csu_programs, tamut_ed4_programs, findMissingClass, set_pricing_cci, set_pricing_au, set_pricing_met, set_pricing_uwlax, set_pricing_csu, set_pricing_tamut_ed4, set_commission, set_pricing_column


def lambda_handler(event, context):
    # month = event['month']
    # current_month = event['current_month']
    month = event['queryStringParameters']['month']
    current_month = event['queryStringParameters']['current_month']
    dbx = dropbox.Dropbox(os.environ['DROPBOX_CODE'])
    s3 = boto3.resource(
        's3', aws_access_key_id=os.environ['AWSAccessKeyId'], aws_secret_access_key=os.environ['AWSSecretKey'])
    pete = "/ECA Back Office/Pete's Backup/MILTARY/PETE ALL 3 SPREADSHEETS MYCAA FOR STACEY AND LISA/MAIN ENROLLMENT FOLDER/SPREADSHEETS/students mycaa FINAL-TODAY.xlsx"

    # Run the Dowload Lambda to get sheets from Dropbox to S3
    # inputForInvoker = {"month": month}
    # client = boto3.client('lambda')
    # response = client.invoke(FunctionName='arn:aws:lambda:us-east-1:077679469516:function:downloadSchoolSpreadsheets', InvocationType='RequestResponse', Payload=json.dumps(inputForInvoker))
    # print('finished invoke')
    # dowload files from S3 Bucket to /tmp dir

    try:
        s3.Bucket('jobautomation').download_file(
            f'{month} 2020.xlsx', f'/tmp/{month} 2020.xlsx')
        s3.Bucket('jobautomation').download_file(
            'students mycaa FINAL-TODAY.xlsx', '/tmp/students mycaa FINAL-TODAY.xlsx')
        s3.Bucket('jobautomation').download_file(
            'Jon weekly email list.xlsx', '/tmp/Jon weekly email list.xlsx')
        print('downloaded to /tmp')
    except botocore.exceptions.ClientError as e:
        if e.response['Error']['Code'] == "404":
            print("The object does not exist.")
        else:
            raise

    monthly_spreadsheet = f"/tmp/{month} 2020.xlsx"
    jon_email_workbook = "/tmp/Jon weekly email list.xlsx"
    pete_spreadsheet = "/tmp/students mycaa FINAL-TODAY.xlsx"

    time_start = datetime.now().strftime("%H:%M:%S")
    print("Starting workbook load at " + time_start)
    wb1 = xl.load_workbook(pete_spreadsheet)
    wb2 = xl.load_workbook(monthly_spreadsheet)
    wb3 = xl.load_workbook(jon_email_workbook)
    time_end = datetime.now().strftime("%H:%M:%S")
    print('Load finished at ' + time_end)

    auburn = wb1.worksheets[0]
    clemson = wb1.worksheets[1]
    csu = wb1.worksheets[2]
    lsu = wb1.worksheets[3]
    msu = wb1.worksheets[4]
    unh = wb1.worksheets[5]
    tamu = wb1.worksheets[7]
    wku = wb1.worksheets[8]
    utep = wb1.worksheets[9]
    uwlax = wb1.worksheets[10]
    desu = wb1.worksheets[12]
    tamiu = wb1.worksheets[14]
    utep = wb1.worksheets[9]
    wtamu = wb1.worksheets[15]
    monthly = wb2.worksheets[0]
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
                f'Smart lookup finished. {num}% that {wrong} is {name}')
            return dictionary[name]
        else:
            print("Smart lookup finished. Nothing really seems to match")
            return dictionary[name]

    def nameCleaner(x):
        if '-LAPTOP' in x:
            name, laptop = x.split('-LAPTOP')
            return name
        elif 'LAPTOP' in x:
            name, laptop = x.split('LAPTOP')
            return name
        else:
            return x

    def findDoubleStudent():
        name_list = dict({})
        num = 1
        students = 0
        for i in range(3, 150):
            name = monthly.cell(row=i, column=4).value
            if name == None:
                pass
            else:
                newName = nameCleaner(name)
                if newName in name_list.values():
                    print(f'{newName} is a double name')
                    students += 1
                else:
                    name_list[num] = nameCleaner(name)
                    num += 1
        if students < 1:
            print('no double students found')
            return 'no double students found'
        else:
            return "There are double students"

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
                    monthly.cell(
                        row=num, column=11).value = last_invoice_number+1
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
            if schoolString == 'TAMIU':
                name = school.cell(row=i, column=8).value
            else:
                name = school.cell(row=i, column=9).value
            last_number_row = num - 1
            if c != None and c.strftime('%Y') == '2020' and c.strftime('%m') == current_month:
                if findName(name) != True:
                    # place invoice number
                    last_invoice_number = monthly.cell(
                        row=last_number_row, column=11).value
                    monthly.cell(
                        row=num, column=11).value = last_invoice_number+1
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

                    if schoolString == 'UTEP' or schoolString == 'WTAMU':
                        address = school.cell(row=i, column=8).value
                    elif schoolString == 'TAMIU':
                        address = school.cell(row=i, column=9).value
                    else:
                        address = school.cell(row=i, column=7).value
                    monthly.cell(row=num, column=14).value = address

                    if schoolString == 'UTEP' or schoolString == 'TAMIU' or schoolString == 'WTAMU':
                        email = school.cell(row=i, column=7).value
                    else:
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

                    # checks the rep column for school
                    if schoolString == 'UNH':
                        rep = school.cell(row=i, column=13).value
                    else:
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

                    if vender == 'ED4O' and schoolString == 'TAMU' or vender == 'ED40' and schoolString == 'TAMU':
                        monthly.cell(row=num, column=9).value = 'TAMU ED4'
                        monthly.cell(row=num, column=set_pricing_column(
                            'TAMU')).value = set_pricing_tamut_ed4(course)
                    elif vender == 'ED4O' and schoolString == 'DESU':
                        monthly.cell(row=num, column=9).value = 'DESU ED4'
                        monthly.cell(row=num, column=set_pricing_column(
                            'DESU')).value = set_pricing_met(course)
                    elif schoolString == 'CSU':
                        monthly.cell(row=num, column=set_pricing_column(
                            schoolString)).value = set_pricing_csu(course)
                    elif schoolString == 'UWLAX':
                        monthly.cell(row=num, column=set_pricing_column(
                            schoolString)).value = set_pricing_uwlax(course)
                    elif vender == 'Pete Medd' or vender == 'PETE MEDD':
                        monthly.cell(row=num, column=9).value = 'TAMU M'
                        monthly.cell(row=num, column=set_pricing_column(
                            schoolString)).value = set_pricing_cci(course)
                    else:
                        monthly.cell(row=num, column=set_pricing_column(
                            schoolString)).value = set_pricing_cci(course)

                    num += 1
                    num1 += 1

    def pete_commission():
        num = 0
        for cell in monthly["F"]:
            if cell.value == "pete code lead" or cell.value == 'pete':
                num += 1
        return num

    def runProgram(date):
        try:
            start = findNextCell()
            auburn_students(date)
            school_tab(date, clemson, 'CLEM', 450)
            school_tab(date, csu, 'CSU', 96)
            school_tab(date, lsu, 'LSU', 74)
            school_tab(date, msu, 'MSU', 450)
            school_tab(date, unh, 'UNH', 26)
            school_tab(date, tamu, 'TAMU', 50)
            school_tab(date, wku, 'WKU', 257)
            school_tab(date, uwlax, 'UWLAX', 246)
            school_tab(date, desu, 'DESU', 11)
            school_tab(date, tamiu, 'TAMIU', 9)
            school_tab(date, utep, 'UTEP', 9)
            school_tab(date, wtamu, 'WTAMU', 10)
            wb2.save(monthly_spreadsheet)
            wb3.save(jon_email_workbook)
            end = findNextCell()
            total = end-start
            print("All Done Transferring Students!")
            print(f"{total} were transferred")
            doubles = findDoubleStudent()
            return total, doubles
        except:
            return 'Something went wrong', ''
            print('Something went wrong :-(')

    students, double_students = runProgram(current_month)
    # s3.Bucket('jobautomation').upload_file(f'/tmp/{month} 2020.xlsx', f'{month} 2020(updated).xlsx')
    # s3.Bucket('jobautomation').upload_file(jon_email_workbook, 'Jon weekly email list.xlsx')

    overwrite = True
    mode = (dropbox.files.WriteMode.overwrite
            if overwrite
            else dropbox.files.WriteMode.add)

    with open(f'/tmp/{month} 2020.xlsx', 'rb') as f:
        dbx.files_upload(
            f.read(), f'/ECA Back Office/JON/{month} 2020(updated).xlsx', mode, mute=True)
        f.close()
    with open(jon_email_workbook, 'rb') as f:
        dbx.files_upload(
            f.read(), '/ECA Back Office/JON/Jon weekly email list.xlsx', mode, mute=True)
        f.close()

    # response = client.invoke(FunctionName='arn:aws:lambda:us-east-1:077679469516:function:downloadSchoolSpreadsheets', InvocationType='RequestResponse', Payload=json.dumps(inputForInvoker))

    return {
        'statusCode': 200,
        "headers": {'Access-Control-Allow-Origin': '*'},
        'body': json.dumps({'Message': 'Finished Transferring MYCAA Students to Monthly Spreadsheet', 'Transferred_Students': students, 'Number_of_Double_Names': double_students})
    }
