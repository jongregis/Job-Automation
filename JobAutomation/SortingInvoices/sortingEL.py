import os
import shutil

# Create a folder to store invoices
desktop = '/Users/jongregis/Desktop/E-Learning Invoices'
os.mkdir(desktop)

# Change dir to where invoices are stored
os.chdir('/Users/jongregis/Python/JobAutomation/practice invoices/E-Learning Invoices')

# Go through files and only grab CCI Invoices
for f in os.listdir():
    shutil.move(f, desktop)

# Create zip folder from the CCI Invoices
os.chdir('/Users/jongregis/Desktop')
shutil.make_archive('E-Learning Invoices', 'zip', desktop)

correct = input('Does everything look correct? Y/N: ')

if correct == 'y':
    mycaa = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Invoices/Invoices"
    Elearning = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Invoices/ILT Invoices"
    os.chdir(desktop)
    for f in os.listdir():
        if f == '.DS_Store':
            continue
        if 'BROWARD' in f or 'SCHREINER' in f or 'MN State' or 'East MS' in f:
            shutil.move(f, Elearning)
