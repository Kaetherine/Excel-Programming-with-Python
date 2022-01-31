from faker import Faker
import xlwings as xw
import random

'''example use case: tracking of deliverys/ services to
support creation of revenue surplus statement'''

def main():
	'''opens front-end app and prompts for input'''
	pass

@xw.func
def print_delivery_note():
	'''creates and saves delivery note to directory of choice'''
	pass

@xw.func
def print_bill():
	'''
	-creates and saves bill to directory of choice,
	-contains delivery note number,
	-registers billing information in excel overview
	'''
	pass

@xw.func
def print_billing_register():
	'''saves overview of billing register to read'''
	pass

@xw.func
def add_artikle():
	'''adds artikle to assortment'''
	pass

@xw.func
def add_customer():
	'''adds customer to sheet "customers"'''
	pass

@xw.func
def payment_update():
	'''adds single payments and updates the payment status'''
	pass

@xw.func
def find(data):
	'''finds customers, bills and delivery notes in spreadsheets'''
	if data == 'customers':
		pass
	elif data == 'products':
		pass
	elif data == 'delivery note':
		pass

def validate(word):
	'''evaluates input word'''
	pass

def setup():

    fake = Faker(['de'])
    wb = xw.Book.caller()

    customers = wb.sheets[0]
    customers.name = 'customers'

    delivery = wb.sheets.add('delivery note', after='customers')
    products = wb.sheets.add('products', after='delivery note')

    customers_values = [
            'vat_id', 'company', 'street', 'street_no', 'zip', 'city'
            ]
    delivery_values = [
            'delivery_no', 'date', 'delivered to', 'amount', 'amount payed',
            'amount open', 'status'
            ]
    products_values = [
            'product_no', 'product_name', ' price net', 'vat rate',
            'discount', 'reduced', 'sum net'
            ]

    customers.range('A1').value = customers_values
    delivery.range('A1').value = delivery_values
    products.range('A1').value = products_values

    vat_rate = '0,19'
    amount_open = ''
    status = ''
    gross = ''
    delivery_no = ''
    amount_payed = ''
    reduced = ''
    sumnet = ''
    date = ''
    discount = ''
    delivered_to = ''

    i = 2
    for value in range(0, 50):

        street = fake.street_name()
        company = fake.company()
        net = fake.pyfloat(3, 2, min_value=100, max_value=910)
        vat_id  = f'DE{random.randint(391246789, 406546709)}'
        street_no = random.randint(1, 160)
        product_name = fake.first_name()
        city = fake.city()
        postcode = fake.postcode()
        product_no = i-1

        customers.range(f'A{i}').value = [
                vat_id, company, street, street_no, postcode, city
                ]
        delivery.range(f'A{i}').value = [
                delivery_no, date, delivered_to, gross, amount_payed,
                amount_open, status
                ]
        products.range(f'A{i}').value = [
                product_no, product_name, net, vat_rate, discount,
                reduced, sumnet
                ]
        i += 1

if __name__ == "__main__":
    xw.Book("sample_data.xlsm").set_mock_caller()
    main()
