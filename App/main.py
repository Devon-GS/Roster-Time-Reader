from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from PIL import ImageTk,Image
import os
from os import path
import pandas as pd
import attendants as att
import cashiers as cas
import bakery as bak
import builder as build
import format as f
import sqlite3

root = Tk()

# ==============================================================================
# FUNCTIONS
# ==============================================================================
def clean_database():
	con = sqlite3.connect('database/time_Sheet.db')
	c = con.cursor()

	# Clear attendents
	c.execute('DROP TABLE IF EXISTS rosterAttWeekOne')
	c.execute('DROP TABLE IF EXISTS rosterAttWeekTwo')

	# Clear cashiers
	c.execute('DROP TABLE IF EXISTS rosterCashierWeekOne')
	c.execute('DROP TABLE IF EXISTS rosterCashierWeekTwo')
	
	# Clear bakers
	c.execute('DROP TABLE IF EXISTS rosterBakerWeekOne')
	c.execute('DROP TABLE IF EXISTS rosterBakerWeekTwo')

	con.commit()
	con.close() 

def sheet_date_checker(attendant, cashier, baker):
	a = attendant[-9:]
	c = cashier[-9:]
	b = baker[-9:]

	if a == c == b:
		return True

def run_calculator():
	try:
		attendant = attendant_selectbox.get()
		cashier = cashier_selectbox.get()
		baker = bakery_selectbox.get()
		
		# Check if data of sheets match 
		if not sheet_date_checker(attendant, cashier, baker):
			ans = messagebox.askyesno(title='Dates Not Matching', message='The dates of the select sheet do not match. Do you want to continue?')
			if not ans:
				raise Exception('Operation Canceled By User')
				
		# Clean database
		clean_database()

		# Remove total time worked.xlsx
		if os.path.exists('Total Time Worked.xlsx'):
			os.remove('Total Time Worked.xlsx')

		# Build attendants Week 1 and Week 2
		if attendant != '':
			at_one = att.attendant_one(attendant)
			at_two = att.attendant_two(attendant)

			build.build_excel('Attendants', at_one[0], at_one[1], week=1)
			build.build_excel('Attendants', at_two[0], at_two[1], week=2)

		# Build cashier Week 1 and Week 2
		if cashier != '':
			cas_one = cas.cashier_one(cashier)
			cas_two = cas.cashier_two(cashier)

			build.build_excel('Cashiers', cas_one[0], cas_one[1], week=1)
			build.build_excel('Cashiers', cas_two[0], cas_two[1], week=2)
		
		# Build baker Week 1 and Week 2
		if baker != '':
			bak_one = bak.baker_one(baker)
			bak_two = bak.baker_two(baker)

			build.build_excel('Bakers', bak_one[0], bak_one[1], week=1)
			build.build_excel('Bakers', bak_two[0], bak_two[1], week=2)

		# Format excel
		if attendant != '':
			f.format('Attendants')
		if cashier != '':
			f.format('Cashiers')

		# Open total time excel
		os.system('start "EXCEL.EXE" "Total Time Worked.xlsx"') 
	except PermissionError:
		messagebox.showerror(title='ERROR', message="Please close 'Total Time Worked.xlsx' and try again!")
	except Exception as error:
		messagebox.showerror(title='ERROR', message=error)

# ==============================================================================
# GUI
# ==============================================================================
	
# Test if files are in location
try:
	file = path.exists('../CASHIERS_ROSTER.xlsx')
	if file == False:
		raise FileNotFoundError
except FileNotFoundError:
	error_label = Label(root, text='File CASHIERS_ROSTER.xls not found')
	error_label.grid(row=1, column=0, sticky=N+E+S+W, pady=(2, 0), padx=(5, 0))
else:
	try:
		file = path.exists('../Attendant_Carwash_Roster.xlsx')
		if file == False:
			raise FileNotFoundError
	except FileNotFoundError:
		error_label = Label(root, text='File Attebdant_Carwash_Roster.xls not found')
		error_label.grid(row=1, column=0, sticky=N+E+S+W, pady=(2, 0), padx=(5, 0))
	else:
		file_a = '../Attendant_Carwash_Roster.xlsx'
		file_sheets_a = pd.ExcelFile(file_a).sheet_names
		
		file_c = '../CASHIERS_ROSTER.xlsx'
		file_sheets_c = pd.ExcelFile(file_c).sheet_names

		db_dir = path.exists('database')
		if db_dir == False:
			os.mkdir('database')

		# LABEL AND COMBOBOX
		# Attendant time
		attendant_selectbox_label = Label(root, text='Select Roster Sheet for Attendant Times')
		attendant_selectbox_label.grid(row=0, column=0, padx=(50,0), pady=(10,0), sticky=W)

		attendant_selectbox = ttk.Combobox(root, value=file_sheets_a)
		attendant_selectbox.grid(row=0, column=1, pady=(10,0), sticky=E)
		
		# Cashier Times
		cashier_selectbox_label = Label(root, text='Select Roster Sheet for Cashier Times')
		cashier_selectbox_label.grid(row=1, column=0, padx=(50,0), pady=(10,0), sticky=W)

		cashier_selectbox = ttk.Combobox(root, value=file_sheets_c)
		cashier_selectbox.grid(row=1, column=1, pady=(10,0), sticky=E)

		# Bakery Times
		bakery_selectbox_label = Label(root, text='Select Roster Sheet for Bakery Times')
		bakery_selectbox_label.grid(row=2, column=0, padx=(50,0), pady=(10,0), sticky=W)

		bakery_selectbox = ttk.Combobox(root, value=file_sheets_c)
		bakery_selectbox.grid(row=2, column=1, pady=(10,0), sticky=E)

		# Run Time Calculator
		run_button = ttk.Button(root, text="Run Calculator", command=run_calculator)
		run_button.grid(row=3, column=0, columnspa=2, padx=(50,0), pady=(10,0), sticky=W+E)

# Window layout
root.title('Calculate Roster Time')
root.iconbitmap('icons/time.ico')
# root.columnconfigure(0, minsize=200)
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
app_width = 450 
app_height = 150
x = (screen_width / 2) - (app_width / 2)
y = (screen_height / 2) - (app_height / 2)
root.geometry(f'{app_width}x{app_height}+{int(x)}+{int(y)}')
root.mainloop()