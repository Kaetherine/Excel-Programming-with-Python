from faker import Faker
import xlwings as xw

fake = Faker()
wb = xw.Book()
wb.save('Rechnung-u_Lieferscheinerst.xlsx')

wb = xw.Book('Rechnung-u_Lieferscheinerst.xlsx')
customers = wb.sheets[0]
customers.name = 'customers'
suppliers = wb.sheets.add('suppliers', after='customers')

customers.range('A1').value = [['0', fake.company(), fake.name()], ['0', fake.company(), fake.name()]]
customers.range('A1').expand().value
