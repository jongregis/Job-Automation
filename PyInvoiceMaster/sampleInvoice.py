from datetime import datetime, date
from .models import InvoiceInfo, ServiceProviderInfo, ClientInfo, Item, Transaction
from .template import SimpleInvoice
from dateutil.relativedelta import *


def create_invoice(start_date, name, description, tuition, percentage, school, invoice_number):

    doc = SimpleInvoice(
        '/Users/jongregis/Python/JobAutomation/practice invoices/MYCAA Invoices/{} {} ({}).pdf'.format(invoice_number, name, school))

    # Paid stamp, optional
    doc.is_paid = 'ECA'

    current_date = datetime.now().strftime('%m'+'/'+'%d'+'/'+'%-y')
    due_date = datetime.now() + relativedelta(months=+1)
    due_date = due_date.strftime('%m'+'/'+'%d'+'/'+'%-y')

    doc.invoice_info = InvoiceInfo(
        invoice_number, current_date, due_date)

    # Service Provider Info, optional
    doc.service_provider_info = ServiceProviderInfo(
        name='PSF International LLC',
        street='1257 Water St',
        city='Wrightsville',
        state='PA',
        # country='My Country',
        post_code='17368',

        # vat_tax_number='Vat/Tax number'
    )

    # Client info, optional
    doc.client_info = ClientInfo(name=name, school=school)

    # Calculate Tuition
    if tuition == 2962.5:
        tuition = 3950
        percentage = '25%'
    elif tuition == 2999.25:
        tuition = 3999
        percentage = '25%'
    elif tuition == 2849.25:
        tuition = 3799
        percentage = '25%'
    elif tuition == 3000:
        tuition = 4000
        percentage = '25%'
    elif school == 'UWLAX' and tuition == 2765:
        percentage = '30%'
        tuition = 3950
    elif school == 'UWLAX' and tuition == 2799.30:
        percentage = '30%'
        tuition = 3999
    elif school == 'UWLAX' and tuition == 2695:
        percentage = '30%'
        tuition = 3850
    elif school == 'UWLAX' and tuition == 2100:
        percentage = '30%'
        tuition = 3000
    elif school == 'CSU':
        percentage = '25%'
    elif tuition == 3061.25:
        tuition = '3950'
        percentage = '22.5%'
    elif tuition == 3099.23:
        tuition = '3999'
        percentage = '22.5%'
    elif tuition == 2944.23:
        tuition = '3799'
        percentage = '22.5%'
    elif tuition == 4874.25:
        percentage = '25%'
        tuition = 6499
    elif tuition == 5656.73:
        percentage = '22.5%'
        tuition = 7299
    elif tuition == 1704.23:
        percentage = '22.5%'
        tuition = 2199
    elif tuition == 1649.25:
        percentage = '25%'
        tuition = 2199
    elif tuition == 2250:
        percentage = '25%'
        tuition = 3000
    elif tuition == 2625:
        percentage = '25%'
        tuition = 3500
    elif tuition == 2518.75:
        percentage = '22.5%'
        tuition = 3250
    elif tuition == 2437.50:
        percentage = '25%'
        tuition = 3250
    elif tuition == 3100:
        percentage = '22.5%'
        tuition = 4000

    # Add Item
    doc.add_item(Item(start_date, description, tuition, percentage))

    # Add Laptop if
    if int(tuition) < 1650 or school != 'UWLAX' and description == "dental assisting certification" or description == "dental assisting" and school != 'UWLAX':
        if "PP" not in str(invoice_number):
            doc.add_item(Item('', 'Laptop', '90', ''))

    # Tax rate, optional
    # doc.set_item_tax_rate(20)  # 20%

    # Optional
    doc.set_bottom_tip(
        "Email: paul.fears@psfinternational.com<br /><strong>Make All Checks Payable To PSF International, LLC</strong><br/>Thank You For Your Business!")

    doc.finish()


def create_invoice_PSF(start_date, name, description, tuition, percentage, school, invoice_number):

    doc = SimpleInvoice(
        '/Users/jongregis/Python/JobAutomation/practice invoices/TAMU ED4/{} {} ({}).pdf'.format(invoice_number, name, school))

    # Paid stamp, optional
    doc.is_paid = 'PSF'

    current_date = datetime.now().strftime('%m'+'/'+'%d'+'/'+'%-y')
    due_date = datetime.now() + relativedelta(months=+1)
    due_date = due_date.strftime('%m'+'/'+'%d'+'/'+'%-y')

    doc.invoice_info = InvoiceInfo(
        invoice_number, current_date, due_date)

    # Service Provider Info, optional
    doc.service_provider_info = ServiceProviderInfo(
        name='PSF International LLC',
        street='1257 Water St',
        city='Wrightsville',
        state='PA',
        # country='My Country',
        post_code='17368',

        # vat_tax_number='Vat/Tax number'
    )

    # Client info, optional
    doc.client_info = ClientInfo(name=name, school=school)

    # Calculate Tuition
    if tuition == 2962.5:
        tuition = 3950
        percentage = '25%'
    elif tuition == 2999.25:
        tuition = 3999
        percentage = '25%'
    elif tuition == 2849.25:
        tuition = 3799
        percentage = '25%'
    elif tuition == 3000:
        tuition = 4000
        percentage = '25%'
    elif school == 'TAMUT' and tuition == 5499:
        tuition = 6499
    elif school == 'CSU':
        percentage = '25%'
    elif tuition == 3061.25:
        tuition = '3950'
        percentage = '22.5%'
    elif tuition == 3099.23:
        tuition = '3999'
        percentage = '22.5%'
    elif tuition == 2944.23:
        tuition = '3799'
        percentage = '22.5%'
    elif tuition == 4874.25:
        percentage = '25%'
        tuition = 6499
    elif tuition == 5656.73:
        percentage = '22.5%'
        tuition = 7299
    elif tuition == 1704.23:
        percentage = '22.5%'
        tuition = 2199
    elif tuition == 1649.25:
        percentage = '25%'
        tuition = 2199
    elif tuition == 2250:
        percentage = '25%'
        tuition = 3000
    elif tuition == 2625:
        percentage = '25%'
        tuition = 3500
    elif tuition == 2518.75:
        percentage = '22.5%'
        tuition = 3250
    # Ed4 TAMUT
    elif tuition == 2999 and school == 'TAMUT':
        tuition = 3999

    elif tuition == 2750 and school == 'TAMUT':
        tuition = 3750

    # Add Item
    doc.add_item(Item(start_date, description, tuition, percentage))

    # Add Laptop if
    # if int(tuition) < 1650 or school != 'UWLAX' and description == "dental assisting certification" or description == "dental assisting" and school != 'UWLAX':
    #     doc.add_item(Item('', 'Laptop', '90', ''))

    # Tax rate, optional
    # doc.set_item_tax_rate(20)  # 20%

    # Optional
    doc.set_bottom_tip(
        "Email: paul.fears@psfinternational.com<br /><strong>Make All Checks Payable To PSF International, LLC</strong><br/>Thank You For Your Business!")

    doc.finish()


def create_invoice_Ed4(start_date, name, description, tuition, percentage, school, invoice_number):

    doc = SimpleInvoice(
        '/Users/jongregis/Python/JobAutomation/practice invoices/MYCAA Invoices/{} {} ({}).pdf'.format(invoice_number, name, school))

    # Paid stamp, optional
    doc.is_paid = True

    current_date = datetime.now().strftime('%m'+'/'+'%d'+'/'+'%-y')
    due_date = datetime.now() + relativedelta(months=+1)
    due_date = due_date.strftime('%m'+'/'+'%d'+'/'+'%-y')

    doc.invoice_info = InvoiceInfo(
        invoice_number, current_date, due_date)

    # Service Provider Info, optional
    doc.service_provider_info = ServiceProviderInfo(
        name='771 Cool Creek Rd LLC',
        street='1257 Water St',
        city='Wrightsville',
        state='PA',
        # country='My Country',
        post_code='17368',

        # vat_tax_number='Vat/Tax number'
    )

    # Client info, optional
    doc.client_info = ClientInfo(name=name, school=school)

    # Calculate Tuition
    if tuition == 2962.5:
        tuition = 3950
        percentage = '25%'
    elif tuition == 2999.25:
        tuition = 3999
        percentage = '25%'
    elif tuition == 2849.25:
        tuition = 3799
        percentage = '25%'
    elif tuition == 3000:
        tuition = 4000
        percentage = '25%'
    elif school == 'UWLAX' and tuition == 2765:
        percentage = '30%'
        tuition = 3950
    elif school == 'UWLAX' and tuition == 2799.30:
        percentage = '30%'
        tuition = 3999
    elif school == 'UWLAX' and tuition == 2695:
        percentage = '30%'
        tuition = 3850
    elif school == 'UWLAX' and tuition == 2100:
        percentage = '30%'
        tuition = 3000
    elif school == 'CSU':
        percentage = '25%'
    elif tuition == 3061.25:
        tuition = '3950'
        percentage = '22.5%'
    elif tuition == 3099.23:
        tuition = '3999'
        percentage = '22.5%'
    elif tuition == 2944.23:
        tuition = '3799'
        percentage = '22.5%'
    elif tuition == 4874.25:
        percentage = '25%'
        tuition = 6499
    elif tuition == 5656.73:
        percentage = '22.5%'
        tuition = 7299
    elif tuition == 1704.23:
        percentage = '22.5%'
        tuition = 2199
    elif tuition == 1649.25:
        percentage = '25%'
        tuition = 2199
    elif tuition == 2250:
        percentage = '25%'
        tuition = 3000
    elif tuition == 2625:
        percentage = '25%'
        tuition = 3500
    elif tuition == 2518.75:
        percentage = '22.5%'
        tuition = 3250
    elif tuition == 2437.50:
        percentage = '25%'
        tuition = 3250
    elif tuition == 3100:
        percentage = '22.5%'
        tuition = 4000

    # Add Item
    doc.add_item(Item(start_date, description, tuition, percentage))

    # Tax rate, optional
    # doc.set_item_tax_rate(20)  # 20%

    # Optional
    doc.set_bottom_tip(
        "Email: paul.fears@psfinternational.com<br /><strong>Make All Checks Payable To 771 Cool Creek Rd, LLC</strong><br/>Thank You For Your Business!")

    doc.finish()


# create_invoice('2/15/2021', 'Tatiana Smith',
#                'professional bookkeeping with quickbooks', 2849.25, '', 'MET', '9903')
