from datetime import datetime, date
from .models import InvoiceInfo, ServiceProviderInfo, ClientInfo, Item, Transaction
from .templateCoolCreek import SimpleInvoiceCoolCreek
from dateutil.relativedelta import *


def create_invoice_ELearning(start_date, name, description, tuition, percentage, school, invoice_number):

    doc = SimpleInvoiceCoolCreek(
        '/Users/jongregis/Python/JobAutomation/practice invoices/E-Learning Invoices/E-L {} {} ({}).pdf'.format(invoice_number, name, school))

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

    if tuition == 2275:
        tuition = 3250
        percentage = '30%'
    # ECDC
    elif tuition == 700:
        percentage = '30%'
        tuition = 1000
    elif tuition == 1225:
        percentage = '30%'
        tuition = 1750

    # Calculate Tuition

    # Add Item
    doc.add_item(Item(start_date, description, tuition, percentage))

    # Tax rate, optional
    # doc.set_item_tax_rate(20)  # 20%

    # Optional
    doc.set_bottom_tip(
        "Email: paul.fears@psfinternational.com<br /><strong>Make All Checks Payable To 771 Cool Creek Rd, LLC</strong><br/>Thank You For Your Business!")

    doc.finish()


create_invoice_ELearning('12/28/2020', 'Hanieh Mansouri',
                         'Certified Professional in Distribution and Warehousing', 1225, '', 'ECDC', 32)
