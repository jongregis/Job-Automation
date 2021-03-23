from datetime import datetime, date
from ecdc_models import InvoiceInfo, ServiceProviderInfo, ClientInfo, Item, Transaction
from template_ecdc import SimpleInvoiceECDC
from dateutil.relativedelta import *


def create_invoice_ECDC(start_date, name, description, tuition, percentage, school, invoice_number):

    doc = SimpleInvoiceECDC(
        f'/Users/jongregis/Python/JobAutomation/practice invoices/E-Learning Invoices/E-L {invoice_number} ({school}).pdf')

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
    doc.client_info = ClientInfo(
        name='839 North Pine St Wilington DE, 19801', school="Eastside Career Development Center", email="Food Services & Operations Cert Program w/ Customer Service Fundamentals")

    if tuition == 2275:
        tuition = 3250
        percentage = '30%'
    # ECDC
    elif tuition == 700:
        percentage = '30%'
        tuition = 1000

    name_list = [
        'Patricia Coverdale'
    ]
    # food services [1415,70,15,1500] Food Services & Operations Cert Program w/ Customer Service Fundamentals
    # hvac [3480,70,200,3750] HVAC Technician Cert Program
    # Calculate Tuition
    for x in name_list:
        doc.add_item(Item(x, '1415', 70, '15'))
    # doc.add_item(Item('Paulette Vickie Lemon', '1255', 70, '425'))

    # Add Item

    # Tax rate, optional
    # doc.set_item_tax_rate(20)  # 20%

    # Optional
    doc.set_bottom_tip(
        "Email: paul.fears@psfinternational.com<br /><strong>Make All Checks Payable To 771 Cool Creek Rd, LLC</strong><br/>Thank You For Your Business!")

    doc.finish()


create_invoice_ECDC('', 'Eastside Career Development Center',
                    'introduction to trades', 700, '', 'ECDC', '208')
# emails to send to: esr@centralbaptistcdc.org, sbordrick@centralbaptistcdc.org
