from faker import Faker
import xlwings as xw
import random

fake = Faker(['de'])
wb = xw.Book()
wb.save('Rechnung-u_Lieferscheinerstellung.xlsx')
wb = xw.Book('Rechnung-u_Lieferscheinerstellung.xlsx')

customers = wb.sheets[0]
customers.name = 'customers'

delivery = wb.sheets.add('delivery note', after='customers')
products = wb.sheets.add('products', after='delivery note')

customers_values = [
        'vat_id', 'company', 'street', 'street_no', 'zip', 'city'
        ]
delivery_values = [
        'delivery_no', 'date', 'company', 'amount', 'amount payed',
        'amount open', 'status'
        ]
products_values = [
        'product_no', 'product_name', 'net', 'vat rate',
        'discount', 'reduced', 'sum net'
        ]

customers.range('A1').value = customers_values
delivery.range('A1').value = delivery_values
products.range('A1').value = products_values

product_no = '**VBA'
vat_rate = '0,19'
amount_open = '**VBA'
status = ''
amount = 'python'
delivery_no = '**VBA'
amount_payed = '**VBA'
reduced = '**VBA'
sumnet = '**VBA'
date = ''

i = 2
for value in range(0, 50):

    street = fake.street_name()
    company = fake.company()
    net = fake.pyfloat(3, 3, min_value=100, max_value=910)
    vat_id  = f'DE{random.randint(391246789, 406546709)}'
    street_no = random.randint(1, 160)
    discount = random.randint(1,30)/100
    product_name = fake.first_name()
    city = fake.city()
    postcode = fake.postcode()

    customers.range(f'A{i}').value = [vat_id, company, street, street_no, postcode, city]
    delivery.range(f'A{i}').value = [delivery_no, date, 'get_company()', amount, amount_payed, amount_open, status]
    products.range(f'A{i}').value = [product_no, product_name, net, vat_rate, discount, reduced, sumnet]
    i += 1


