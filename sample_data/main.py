from tkinter import *
from tkinter import ttk
import operations


def main():
	'''opens gui app and prompts for input'''
	operation = operations.Operations()
	root = Tk()
	frame = ttk.Frame(root, padding=32)
	frame.grid(columnspan=3, rowspan=100)
	# prompt = ttk.Label(frame, text="Pick a task.").grid(column=0, row=0)

	btn_exp_bill = ttk.Button(
		frame, text="PDF Rechnung",
		command=operation.bill_to_pdf).grid(
			column=1, row=2, ipady=8, ipadx=32
			)
	btn_exp_bill_register = ttk.Button(
		frame, text="PDF Rechnungsregister",
		command = operation.billregister_to_pdf).grid(
			column=1, row=3, ipady=8, ipadx=10
			)
	btn_delete_data = ttk.Button(
		frame, text="Daten löschen",
		command=operation.delete_data).grid(
			column=1, row=4, ipady=8, ipadx=34
			)
	btn_imp_sample_data = ttk.Button(
		frame, text="Importiere Beispieldaten",
		command=operation.setup_sampledata).grid(
			column=1, row=5, ipady=8, ipadx=8
			)
	btn_quit = ttk.Button(
		frame, text="Schließen",
		command=lambda:[operation.quit(), root.quit()]).grid(
			column=1, row=6, ipady=8, ipadx=39
			)
	root.mainloop()

if __name__ == "__main__":
	main()
