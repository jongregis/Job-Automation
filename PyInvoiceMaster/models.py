from __future__ import unicode_literals
from decimal import Decimal


class PDFInfo(object):
    """
    PDF Properties
    """

    def __init__(self, title=None, author=None, subject=None):
        """
        PDF Properties
        :param title: PDF title
        :type title: str or unicode
        :param author: PDF author
        :type author: str or unicode
        :param subject: PDF subject
        :type subject: str or unicode
        """
        self.title = title
        self.author = author
        self.subject = subject
        self.creator = 'PSF International, LLC'


class InvoiceInfo(object):
    """
    Invoice information
    """

    def __init__(self, invoice_id=None, invoice_datetime=None, due_datetime=None):
        """
        Invoice info
        :param invoice_id: Invoice id
        :type invoice_id: int or str or unicode or None
        :param invoice_datetime: Invoice create datetime
        :type invoice_datetime: str or unicode or datetime or date
        :param due_datetime: Invoice due datetime
        :type due_datetime: str or unicode or datetime or date
        """
        self.invoice_id = invoice_id
        self.invoice_datetime = invoice_datetime
        self.due_datetime = due_datetime


class AddressInfo(object):
    def __init__(self, name=None, street=None, city=None, state=None, country=None, post_code=None, phone=None):
        """
        :type name: str or unicode or None
        :type street: str or unicode or None
        :type city: str or unicode or None
        :type state: str or unicode or None
        :type country: str or unicode or None
        :type post_code: str or unicode or int or None
        """
        self.name = name
        self.street = street
        self.city = city
        self.state = state
        self.country = country
        self.post_code = post_code
        self.phone = phone


class ServiceProviderInfo(AddressInfo):
    """
    Service provider/Merchant information
    """

    def __init__(self, name=None, street=None, city=None, state=None, country=None, post_code=None, phone=None,
                 vat_tax_number=None):
        """
        :type name: str or unicode or None
        :type street: str or unicode or None
        :type city: str or unicode or None
        :type state: str or unicode or None
        :type country: str or unicode or None
        :type post_code: str or unicode or None
        :type vat_tax_number: str or unicode or int or None
        """
        super(ServiceProviderInfo, self).__init__(
            name, street, city, state, country, post_code, phone)
        self.vat_tax_number = vat_tax_number


class ClientInfo(AddressInfo):
    """
    Client/Custom information
    """

    def __init__(self, name=None, street=None, city=None, state=None, country=None, post_code=None,
                 email=None, client_id=None, student_name=None, school=None):
        """
        :type name: str or unicode or None
        :type street: str or unicode or None
        :type city: str or unicode or None
        :type state: str or unicode or None
        :type country: str or unicode or None
        :type post_code: str or unicode or None
        :type email: str or unicode or None
        :type client_id: str or unicode or int or None
        """
        super(ClientInfo, self).__init__(
            name, street, city, state, country, post_code)
        self.email = email
        self.client_id = client_id
        self.student_name = student_name
        self.school = school


class Item(object):
    """
    Product/Item information
    """

    def __init__(self, start_date, description, unit_price, percentage):
        """
        Item modal init
        :param name: Item name
        :type name: str or unicode or int
        :param description: Item detail
        :type description: str or unicode or int
        :param units: Amount
        :type units: int or str or unicode
        :param unit_price: Unit price
        :type unit_price: Decimal or str or unicode or int or float
        :return:
        """
        self.percentage = percentage
        self.description = description
        self.start_date = start_date
        self.unit_price = unit_price

    @property
    def amount(self):
        if self.unit_price == 3950 and self.percentage == '25%':
            return Decimal(str(2962.5))
        elif self.unit_price == 3999 and self.percentage == '25%':
            return Decimal(str(2999.25))
        elif self.unit_price == 3799 and self.percentage == '25%':
            return Decimal(str(2849.25))
        elif self.unit_price == 4000 and self.percentage == '25%':
            return Decimal(str(3000))
        elif self.unit_price == '3950':
            return Decimal(str(3061.25))
        elif self.unit_price == '3999':
            return Decimal(str(3099.23))
        elif self.unit_price == '3799':
            return Decimal(str(2944.23))
        elif self.percentage == '30%' and self.unit_price == 3950:
            return Decimal(str(2765))
        elif self.percentage == '30%' and self.unit_price == 3999:
            return Decimal(str(2799.30))
        elif self.percentage == '30%' and self.unit_price == 3850:
            return Decimal(str(2695))
        elif self.unit_price == 6499:
            return Decimal(str(4874.25))
        elif self.unit_price == 7299 and self.percentage == '22.5%':
            return Decimal(str(5656.73))
        elif self.unit_price == 7299 and self.percentage == '25%':
            return Decimal(str(5474.25))

        return Decimal(str(self.unit_price))


class Transaction(object):
    """
    Transaction information
    """

    def __init__(self, gateway, transaction_id, transaction_datetime, amount):
        """
        :param gateway: Payment gateway like Paypal, Stripe etc.
        :type gateway: str or unicode
        :param transaction_id:
        :type transaction_id: int or str or unicode
        :param transaction_datetime:
        :type transaction_datetime: date or datetime or str or unicode
        :param amount: $$
        :type amount: int or float or str or unicode
        :return:
        """
        self.gateway = gateway
        self.transaction_id = transaction_id
        self.transaction_datetime = transaction_datetime
        self.amount = amount
