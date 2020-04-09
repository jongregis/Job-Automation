from datetime import datetime, date
from .models import InvoiceInfo, ServiceProviderInfo, ClientInfo, Item, Transaction
from .template import SimpleInvoice
from dateutil.relativedelta import *


def create_invoice(start_date, name, description, tuition, percentage, school, invoice_number):

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
    elif tuition == 5249.25:
        tuition = 6999
        percentage = '25%'
    elif school == 'UWLAX':
        percentage = '30%'
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

    # Add Item
    doc.add_item(Item(start_date, description, tuition, percentage))

    # Add Laptop if
    if int(tuition) < 1650 or school != 'UWLAX' and description == "dental assisting certification" or description == "dental assisting" and school != 'UWLAX':
        doc.add_item(Item('', 'Laptop', '90', ''))

    # Tax rate, optional
    # doc.set_item_tax_rate(20)  # 20%

    # Optional
    doc.set_bottom_tip(
        "Email: paul.fears@psfinternational.com<br /><strong>Make All Checks Payable To PSF International, LLC</strong><br/>Thank You For Your Bussiness!")

    doc.finish()


# create_invoice('03/31/20', 'Lantonette Boothe',
#                'medical billing and coding with medical administrative assistant', 1370, '', 'AU', 8571)
