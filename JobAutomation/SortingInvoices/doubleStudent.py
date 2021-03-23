import openpyxl as xl


def nameCleaner(x):
    if '-LAPTOP' in x:
        name, laptop = x.split('-LAPTOP')
        return name
    elif 'LAPTOP' in x:
        name, laptop = x.split('LAPTOP')
        return name
    else:
        return x


def findDoubleStudent(month):
    monthly_spreadsheet = f"/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Invoices/2021 Enrollment/{month} 2021.xlsx"
    wb1 = xl.load_workbook(monthly_spreadsheet)
    monthly = wb1.worksheets[0]
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
                print(f'\033[1;31m{newName} is a double name \033[0;0m')
                students += 1
            else:
                name_list[num] = nameCleaner(name)
                num += 1
    if students < 1:
        print('\033[1;32mno double students found \033[0;0m')
        return 'no double students found'
    else:
        return "There are double students"


if __name__ == '__main__':
    findDoubleStudent('Jan')
