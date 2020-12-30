import os
import shutil

# Create a folder to store invoices
desktop = '/Users/jongregis/Desktop/CCI Invoices'
os.mkdir(desktop)

# Change dir to where invoices are stored
os.chdir('/Users/jongregis/Python/JobAutomation/practice invoices/MYCAA Invoices')

# Go through files and only grab CCI Invoices
for f in os.listdir():
    if not 'AU ED4' in f and not 'UWLAX' in f and not 'CSU' in f and not 'MET' in f and not 'TAMU M' in f:
        shutil.move(f, desktop)

# Create zip folder from the CCI Invoices
os.chdir('/Users/jongregis/Desktop')
shutil.make_archive('CCI Invoices', 'zip', desktop)

correct = input('Does everything look correct? Y/N: ')

if correct == 'y':
    mycaa = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Invoices/Invoices"
    Elearning = "/Volumes/SanDisk Extreme SSD/Dropbox (ECA Consulting)/ECA Back Office/Lisa's Backup/Invoices/ILT Invoices"
    os.chdir(desktop)
    for f in os.listdir():
        if f == '.DS_Store':
            continue
        shutil.move(f, mycaa)
