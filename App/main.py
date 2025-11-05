from tkinter import *
from tkinter import ttk
from PIL import ImageTk,Image
import os
from os import path
import pandas as pd
import attendants as att
import cashiers as ca
import bakery as b
import sqlite3

# ############# NOTES #############################################
# If nothing selected do nothing
# Not make so many connection to database
# #################################################################

root = Tk()

def clean_database():
    con = sqlite3.connect('database/time_Sheet.db')
    c = con.cursor()

    c.execute('DROP TABLE IF EXISTS rosterAttWeekOne')
    c.execute('DROP TABLE IF EXISTS rosterAttWeekTwo')

    con.commit()
    con.close() 

def run_calculator():
	attendant = attendant_selectbox.get()
	cashier = cashier_selectbox.get()
	baker = bakery_selectbox.get()

	# Clean database
	clean_database()
	
	
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
app_height = 700
x = (screen_width / 2) - (app_width / 2)
y = (screen_height / 2) - (app_height / 2)
root.geometry(f'{app_width}x{app_height}+{int(x)}+{int(y)}')
root.mainloop()