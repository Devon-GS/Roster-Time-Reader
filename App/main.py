from tkinter import *
from tkinter import ttk
from PIL import ImageTk,Image
import os
from os import path
import pandas as pd
import attendants as a
import cashiers as ca
import bakery as b
import sqlite3

root = Tk()

def attendant_time(selected):
    index = file_sheets_a.index(attendant_selectbox.get())
    wo = a.attendant_one(index)
    wt = a.attendant_two(index)

    conn = sqlite3.connect('database/time_Sheet.db')
    c = conn.cursor()

    c.execute('CREATE TABLE IF NOT EXISTS attendants (name TEXT, standardHours INTEGER, sundayHours INTEGER)')
    c.execute('CREATE TABLE IF NOT EXISTS attendantsWT (name TEXT, standardHours INTEGER, sundayHours INTEGER)')

    c.execute('DELETE FROM attendants')
    c.execute('DELETE FROM attendantsWT')

    # Attendants Times 
    for x in range(len(wo)):
        query = ("INSERT INTO attendants (name, standardHours, sundayHours) VALUES (?, ?, ?)")
        c.execute(query, (wo[x][0], wo[x][1], wo[x][2]))    

    for x in range(len(wt)):
        query = ("INSERT INTO attendantsWT (name, standardHours, sundayHours) VALUES (?, ?, ?)")
        c.execute(query, (wt[x][0], wt[x][1], wt[x][2]))

    c.execute("""
    UPDATE attendants
    SET
    standardHours = standardHours + (SELECT attendantsWT.standardHours
                                            FROM attendantsWT
                                            WHERE attendantsWT.name = attendants.name),
    sundayHours = sundayHours + (SELECT attendantsWT.sundayHours
                                            FROM attendantsWT
                                            WHERE attendantsWT.name = attendants.name)
    WHERE EXISTS (
        SELECT * 
        FROM attendantsWT
        WHERE attendantsWT.name = attendants.name
    )
    """)
    
    c.execute("SELECT * FROM attendants")
    r = c.fetchall()
    
    conn.commit()
    conn.close() 
    
    a_tree = ttk.Treeview(root, height=len(r))
    # Define columns
    a_tree['columns'] = ('Name', 'Standard Hours', 'Sunday Hours', 'Total')
    # formulate columns
    a_tree.column('#0', width=0)
    a_tree.column('Name', anchor=W, width=80)
    a_tree.column('Standard Hours', anchor=CENTER, width=90)
    a_tree.column('Sunday Hours', anchor=CENTER, width=90)
    a_tree.column('Total', anchor=CENTER, width=80)
    # Create Headings
    a_tree.heading('#0', text='')
    a_tree.heading('Name', text='Name', anchor=W)
    a_tree.heading('Standard Hours', text='Standard Hours', anchor=W)
    a_tree.heading('Sunday Hours', text='Sunday Hours', anchor=W)
    a_tree.heading('Total', text='Total', anchor=W)

    for x in range(len(r)):
        name = r[x][0]
        standard = r[x][1] 
        sunday = r[x][2]
        total = r[x][1] + r[x][2]
        a_tree.insert(parent='', index='end', iid=x, text='', values=(name, standard, sunday, total))

    a_tree.grid(row=2, column=0, columnspan=2, padx=(50,0), pady=(10, 0), sticky=W+E)

def cashier_time(selected):
    index = file_sheets_c.index(cashier_selectbox.get())
    cwo = ca.cashier_one(index)
    cwt = ca.cashier_two(index)

    conn = sqlite3.connect('database/time_Sheet.db')
    c = conn.cursor()

    c.execute('CREATE TABLE IF NOT EXISTS cashiers (name TEXT, standardHours INTEGER, sundayHours INTEGER)')
    c.execute('CREATE TABLE IF NOT EXISTS cashiersWT (name TEXT, standardHours INTEGER, sundayHours INTEGER)')

    c.execute('DELETE FROM cashiers')
    c.execute('DELETE FROM cashiersWT')

    # Cashiers Times
    for x in range(len(cwo)):
        query = ("INSERT INTO cashiers (name, standardHours, sundayHours) VALUES (?, ?, ?)")
        c.execute(query, (cwo[x][0], cwo[x][1], cwo[x][2]))   

    for x in range(len(cwt)):
        query = ("INSERT INTO cashiersWT (name, standardHours, sundayHours) VALUES (?, ?, ?)")
        c.execute(query, (cwt[x][0], cwt[x][1], cwt[x][2]))

    c.execute("""
    UPDATE cashiers
    SET
    standardHours = standardHours + (SELECT cashiersWT.standardHours
                                            FROM cashiersWT
                                            WHERE cashiersWT.name = cashiers.name),
    sundayHours = sundayHours + (SELECT cashiersWT.sundayHours
                                            FROM cashiersWT
                                            WHERE cashiersWT.name = cashiers.name)
    WHERE EXISTS (
        SELECT * 
        FROM cashiersWT
        WHERE cashiersWT.name = cashiers.name
    )
    """)
    
    c.execute("SELECT * FROM cashiers")
    r = c.fetchall()
    
    conn.commit()
    conn.close() 
    
    c_tree = ttk.Treeview(root, height=len(r))

    # Define columns
    c_tree['columns'] = ('Name', 'Standard Hours', 'Sunday Hours', 'Total')

    # formulate columns
    c_tree.column('#0', width=0)
    c_tree.column('Name', anchor=W, width=90)
    c_tree.column('Standard Hours', anchor=CENTER, width=90)
    c_tree.column('Sunday Hours', anchor=CENTER, width=90)
    c_tree.column('Total', anchor=CENTER, width=80)

    # Create Headings
    c_tree.heading('#0', text='')
    c_tree.heading('Name', text='Name', anchor=W)
    c_tree.heading('Standard Hours', text='Standard Hours', anchor=W)
    c_tree.heading('Sunday Hours', text='Sunday Hours', anchor=W)
    c_tree.heading('Total', text='Total', anchor=W)

    for x in range(len(r)):
        name = r[x][0]
        standard = r[x][1] 
        sunday = r[x][2]
        total = r[x][1] + r[x][2]
        c_tree.insert(parent='', index='end', iid=x, text='', values=(name, standard, sunday, total))

    c_tree.grid(row=4, column=0, columnspan=2, padx=(50,0), pady=(10, 0), sticky=W+E)

def bakery_time(selected):
    index = file_sheets_c.index(bakery_selectbox.get())
    bwo = b.bakery_one(index)
    bwt = b.bakery_two(index)
    
    conn = sqlite3.connect('database/time_Sheet.db')
    c = conn.cursor()

    c.execute('CREATE TABLE IF NOT EXISTS bakery (name TEXT, standardHours INTEGER, sundayHours INTEGER)')
    c.execute('CREATE TABLE IF NOT EXISTS bakeryWT (name TEXT, standardHours INTEGER, sundayHours INTEGER)')

    c.execute('DELETE FROM bakery')
    c.execute('DELETE FROM bakeryWT')

    # Cashiers Times
    for x in range(len(bwo)):
        query = ("INSERT INTO bakery (name, standardHours, sundayHours) VALUES (?, ?, ?)")
        c.execute(query, (bwo[x][0], bwo[x][1], bwo[x][2]))   

    for x in range(len(bwt)):
        query = ("INSERT INTO bakeryWT (name, standardHours, sundayHours) VALUES (?, ?, ?)")
        c.execute(query, (bwt[x][0], bwt[x][1], bwt[x][2]))

    c.execute("""
    UPDATE bakery
    SET
    standardHours = standardHours + (SELECT bakeryWT.standardHours
                                            FROM bakeryWT
                                            WHERE bakeryWT.name = bakery.name),
    sundayHours = sundayHours + (SELECT bakeryWT.sundayHours
                                            FROM bakeryWT
                                            WHERE bakeryWT.name = bakery.name)
    WHERE EXISTS (
        SELECT * 
        FROM bakeryWT
        WHERE bakeryWT.name = bakery.name
    )
    """)
    
    c.execute("SELECT * FROM bakery")
    r = c.fetchall()
    
    conn.commit()
    conn.close() 
    
    b_tree = ttk.Treeview(root, height=len(r))

    # Define columns
    b_tree['columns'] = ('Name', 'Standard Hours', 'Sunday Hours', 'Total')

    # formulate columns
    b_tree.column('#0', width=0)
    b_tree.column('Name', anchor=W, width=90)
    b_tree.column('Standard Hours', anchor=CENTER, width=90)
    b_tree.column('Sunday Hours', anchor=CENTER, width=90)
    b_tree.column('Total', anchor=CENTER, width=80)

    # Create Headings
    b_tree.heading('#0', text='')
    b_tree.heading('Name', text='Name', anchor=W)
    b_tree.heading('Standard Hours', text='Standard Hours', anchor=W)
    b_tree.heading('Sunday Hours', text='Sunday Hours', anchor=W)
    b_tree.heading('Total', text='Total', anchor=W)

    for x in range(len(r)):
        name = r[x][0]
        standard = r[x][1] 
        sunday = r[x][2]
        total = r[x][1] + r[x][2]
        b_tree.insert(parent='', index='end', iid=x, text='', values=(name, standard, sunday, total))

    b_tree.grid(row=6, column=0, columnspan=2, padx=(50,0), pady=(10, 0), sticky=W+E)

def refresh():
    # Gets selection from drop down
    refresh_attendant = attendant_selectbox.get()
    refresh_cashier = cashier_selectbox.get()
    refresh_bakery = bakery_selectbox.get()

    # Execute functions
    if refresh_attendant == '':
        pass
    else:
        attendant_time(refresh_attendant)
    if refresh_cashier == '':
        pass
    else:
        cashier_time(refresh_cashier)
    if refresh_bakery == '':
        pass
    else:
        bakery_time(refresh_bakery)

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
        # Refresh
        refresh_button = ttk.Button(root, text="Refresh", command=refresh)
        refresh_button.grid(row=0, column=0, columnspa=2, padx=(50,0), pady=(10,0), sticky=W+E)

        # Attendant time
        attendant_selectbox_label = Label(root, text='Select Roster Sheet for Attendant Times')
        attendant_selectbox_label.grid(row=1, column=0, padx=(50,0), pady=(10,0), sticky=W)

        attendant_selectbox = ttk.Combobox(root, value=file_sheets_a)
        attendant_selectbox.bind('<<ComboboxSelected>>', attendant_time)
        attendant_selectbox.grid(row=1, column=1, pady=(10,0), sticky=E)
        
        # Cashier Times
        cashier_selectbox_label = Label(root, text='Select Roster Sheet for Cashier Times')
        cashier_selectbox_label.grid(row=3, column=0, padx=(50,0), pady=(10,0), sticky=W)

        cashier_selectbox = ttk.Combobox(root, value=file_sheets_c)
        cashier_selectbox.bind('<<ComboboxSelected>>', cashier_time)
        cashier_selectbox.grid(row=3, column=1, pady=(10,0), sticky=E)

        # Bakery Times
        bakery_selectbox_label = Label(root, text='Select Roster Sheet for Bakery Times')
        bakery_selectbox_label.grid(row=5, column=0, padx=(50,0), pady=(10,0), sticky=W)

        bakery_selectbox = ttk.Combobox(root, value=file_sheets_c)
        bakery_selectbox.bind('<<ComboboxSelected>>', bakery_time)
        bakery_selectbox.grid(row=5, column=1, pady=(10,0), sticky=E)

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