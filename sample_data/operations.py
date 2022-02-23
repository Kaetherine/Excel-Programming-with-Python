from faker import Faker
import xlwings as xw
from datetime import datetime

class Operations():

	def __init__(self):
		'''opens the excel application and workbook'''
		self.app = xw.App(add_book=False)
		self.wb = self.app.books.open(
			'main.xlsm', read_only=False, ignore_read_only_recommended=True
			)
		self.date = datetime.today().strftime('%Y-%m-%d')
		if self.wb.sheets[1] and self.wb.sheets[2]:
			self.products = self.wb.sheets[0]
			self.customers = self.wb.sheets[1]
			self.bill_overview = self.wb.sheets[2]
			self.products.name = 'Produktkatalog'
			self.customers.name = 'Kundendaten'
			self.bill_overview.name = 'Rechnungsübersicht'
		else:
			self.products = self.wb.sheets[0]
			self.products.name = 'Produktkatalog'
			self.customers = self.wb.sheets.add(
				'Kundendaten', after='Produktkatalog')
			self.bill_overview = self.wb.sheets.add(
				'Rechnungsübersicht', after='Kundendaten')
	
	def name_columns(self):
		'''names all columns to supervise structure for basis data'''
		self.products = self.wb.sheets[0]
		self.products.name = 'Produktkatalog'
		self.customers = self.wb.sheets.add(
			'Kundendaten', after='Produktkatalog')
		self.bill_overview = self.wb.sheets.add(
			'Rechnungsübersicht', after='Kundendaten')
		customers_values = [
				'USt.Id', 'Unternehmensname', 'Straße', 'Hausnr.', 'PLZ',
				'Ort'
				]
		bill_overview_values = [
				'Rechnungsnr.', 'Datum', 'Kundennr.', 'Summe', 'Steuerbetr.',
				'Rechnungsbetr.'
				]
		products_values = ['Artikelnr', 'Artikelname', 'Preis netto', 'Menge']

		self.customers.range('A1').value = customers_values
		self.bill_overview.range('A1').value = bill_overview_values
		self.products.range('A1').value = products_values
	
	def check_len_sheet(self,sheet):
		'''checks length of sheet "Rechnungsübersicht"'''
		sheets = sheet
		row = 2
		end = False
		while end != True:
			cell_check = sheets.range(f'A{row}').value
			if cell_check == None:
				end = True
				break
			else:
				row += 1
		return row

	def reset(self):
		'''prepares application setup to create more bills
		by erasing previous inputs'''
		row = self.check_len_sheet(self.products)
		for value in range(1, row):
			amount = self.products.range(f'D{row}')
			if amount != None:
				amount.clear()
				row -= 1
				continue
			else:
				row -= 1
				continue
	
	def make_breakline(self, cell):
		''''Writes a breakline into the given cell'''
		self.bill.range(f'{cell}').value = (
			'_____________________________________________________________'
			'_______________'
		)

	def quit(self):
		'''saves applied changes and quits excel'''
		self.wb.save()
		self.app.quit()

	def bill_to_pdf(self):
		'''exports and registers billing information in excel overview'''
		seller_data = {
			'vat_id': 'DE12345678', 'company': 'Merve Musterfrau GbR',
			'street': 'Musterstraße', 'str_no': '15', 'zip': '54321',
			'city': 'Berlin', 'phone': '0049 30 2022 123',
			'email': 'buchhaltung@mervem.com', 'bank': 'Berliner Musterbank',
			'iban': 'DE50200001000100008000', 'bic': 'LLO30RLD1CA'
		}
		
		adress_location = self.app.selection.options(numbers=int)
		customer_adress = adress_location.value
		shopping_cart = []

		len_products = self.check_len_sheet(self.products)
		i = 2
		for value in range(0,len_products-1):
			if self.products.range(f'D{i}').value != None:
				item = self.products.range(f'A{i}:D{i}').value
				shopping_cart.append(item)
				i+=1
				continue
			else:
				i+=1
				continue

		customer_data = {}
		customer_data.update({
			'vat_id': customer_adress[0], 'company': customer_adress[1],
			'street': customer_adress[2], 'str_no': customer_adress[3],
			'zip':customer_adress[4], 'city': customer_adress[5]
			})

		after_last_bill = self.check_len_sheet(self.bill_overview)
		customer_hint = customer_data['vat_id'][0:3]
		year = datetime.today().strftime('%Y')
		bill_no = f'{after_last_bill-1}-{year}{customer_hint}'
		self.bill = self.wb.sheets.add(bill_no, after='Rechnungsübersicht')
		self.bill.range('A11').value = [f'Rechnungsnummer {bill_no}']

		s_full_street = f'{seller_data["street"]}, {seller_data["str_no"]}'
		s_full_city = f'{seller_data["zip"]} {seller_data["city"]}'
		sender = f'{seller_data["company"]}, {s_full_street}, {s_full_city}'

		self.bill.range('A1').value = sender
		self.make_breakline('A2')

		full_street = f'{customer_data["street"]}, {customer_data["str_no"]}'
		full_city = f'{customer_data["zip"]} {customer_data["city"]}'
		self.bill.range('A5').value = customer_data["company"]
		self.bill.range('A6').value = full_street
		self.bill.range('A7').value = full_city

		self.bill.range('G9').value = self.date
		
		self.bill.range('A13').value = [
			'Position', 'Artikelnr', 'Artikelbez.', 'Einzelpreis',
			'Menge', None, 'Gesamt'
			]
		self.make_breakline('A14')
		
		net_sum = 0
		pos = 1
		n = 15
		for art in shopping_cart:
			amount = art[-1]
			single_price = float(art[-2])
			total = amount*single_price
			art.insert(0, str(pos))
			art.append(None)
			art.append(f'{total:.2f}')
			net_sum += total
			self.bill.range(f'A{n}').value = art
			n += 1
			pos += 1

		tax_amount = net_sum*0.19
		gross_sum = net_sum + tax_amount
		self.bill.range(f'E{n+1}').value = (
			'________________________________'
		)
		self.bill.range(f'E{n+2}').value = [
			'Summe Netto', None, f'{net_sum:.2f}'
			]
		self.bill.range(f'E{n+3}').value = [
			'19% Umsatzsteuer', None, f'{tax_amount:.2f}'
			]
		self.bill.range(f'E{n+4}').value = [
			'Rechnungsbetrag', None, f'{gross_sum:.2f}'
			]

		self.bill.range('A35').value = 'Satz zum Zahlungstermin.'
		self.bill.range('A37').value = 'Satz zur Aufbewahrungspflicht.'

		self.make_breakline('A39')
		self.bill.range('A40').value = [
			[f'USt.Id.: {seller_data["vat_id"]}'],
			[seller_data['company']],
			[s_full_street],
			[s_full_city],
			[seller_data['phone']],
			[seller_data['email']]
			]
	
		self.bill.range('D41').value = [
			[f'Bank: {seller_data["bank"]}'],
			[f'IBAN: {seller_data["iban"]}'],
			[f'BIC: {seller_data["bic"]}'],
			[f'Kontoinhaberin: {seller_data["company"]}'],
			[f'Verwendungszweck: {bill_no}']
			]
		
		self.bill_overview.range(f'A{after_last_bill}').value = [
			bill_no, self.date, customer_data['vat_id'], f'{net_sum:.2f}',
			f'{tax_amount:.2f}', f'{gross_sum:.2f}'
		]

		self.bill.to_pdf()
		self.reset()
		self.bill.delete()
		self.wb.save()

	def billoverview_to_pdf(self):
		'''saves overview of billing overview to read'''
		self.bill_overview.to_pdf()

	def delete_data(self):
		'''removes all data from all sheets except the headers'''
		len_sheets = len(self.wb.sheets)
		if len_sheets > 1:
			i = len_sheets-1
			while i != 0:
				self.wb.sheets[i].delete()
				self.wb.sheets[0].clear()
				i -= 1
		else:
			self.wb.sheets[0].clear()
		self.wb.sheets[0].name = 'Tabelle1'
		self.name_columns()
		self.wb.save()

	def setup_sampledata(self):
		'''imports sample data to test the application'''
		self.delete_data()
		fake = Faker(['de_DE'])
		i = 2
		for n in range(0, 50):
			street = fake.street_name()
			company = fake.company()
			net = fake.pyfloat(3, 2, min_value=100, max_value=900)
			vat_id  = fake.vat_id()
			street_no = fake.pyint(min_value=1, max_value=160)
			product_name = fake.first_name().upper()
			city = fake.city()
			postcode = fake.postcode()
			product_no = i-1

			self.customers.range(f'A{i}').value = [
					vat_id, company, street, street_no, postcode, city
					]
	
			self.products.range(f'A{i}').value = [
					product_no, product_name, net,
					]
			i += 1
		self.wb.save()




