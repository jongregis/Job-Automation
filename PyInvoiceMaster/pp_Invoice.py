from datetime import datetime, date
from .models import InvoiceInfo, ServiceProviderInfo, ClientInfo, Item, Transaction
from .template import SimpleInvoice
from dateutil.relativedelta import *


def create_invoice_PrivatePay(start_date, name, description, tuition, percentage, school, invoice_number):

    doc = SimpleInvoice(
        f'/Users/jongregis/Python/JobAutomation/practice invoices/PP Invoices/{invoice_number} {name} ({school}).pdf')

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
    if tuition == 5474.25:
        tuition = 7299
        percentage = '25%'
    elif tuition == 5656.73:
        tuition = 7299
        percentage = '22.5%'
    elif tuition == 2962.50:
        tuition = 3950
        percentage = '25%'
    elif tuition == 4874.25:
        tuition = 6499
        percentage = '25%'
    elif tuition == 3187.50:
        tuition = 4250
        percentage = '25%'
    elif tuition == 2250:
        tuition = 3000
        percentage = '25%'
    elif tuition == 2625.50:
        tuition = 3500
        percentage = '25%'
    elif tuition == 2849.25:
        tuition = 3799
        percentage = '25%'

    # Add Item
    doc.add_item(Item(start_date, description, tuition, percentage))

    # Tax rate, optional
    # doc.set_item_tax_rate(20)  # 20%

    # Optional
    doc.set_bottom_tip(
        "Email: paul.fears@psfinternational.com<br /><strong>Make All Checks Payable To PSF International, LLC</strong><br/>Thank You For Your Business!")

    doc.finish()


# create_invoice_PrivatePay('2/12/2021', 'Keyla Vasquez',
#                           'Opthalmic Assistant Specialist', 2962.50, '', 'AU', 'PP60')
