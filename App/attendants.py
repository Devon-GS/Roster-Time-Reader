import pandas as pd
import sqlite3


# ==============================================================================
# ATTENDENTS WEEK 1
# ==============================================================================
def attendant_one(selected):
	file = '../Attendant_Carwash_Roster.xlsx'
	file_sheets = pd.ExcelFile(file).sheet_names

	# Get Columns
	columns = ['idx','ATTENDANTS', 'THURS', 'FRI', 'SAT', 'SUN', 'MON', 'TUE', 'WED']
	
	# Get Selected sheet
	index = file_sheets.index(selected)
	sheet = file_sheets[index]

	# Get Times
	data = pd.read_excel(file, sheet_name=sheet, index_col=0, header=1, usecols=columns)
	data = data.fillna(0)

	# Get Dates
	data_date = pd.read_excel(file, sheet_name=sheet, header=None, nrows=2, usecols='C:I')
	data_date_ex = data_date.loc[0]

	weekone_dates = {}
	for x in data_date_ex:
		weekone_dates[x.strftime("%A")] = x.date().strftime("%d/%m/%y")

	# Get week one from excel roster sheet
	week_one_data = data.loc[0:14]
	week_one = []
	for x in week_one_data.to_numpy().tolist():
		if str(x[0]) != 'nan':
			if x[0] != 0:
				week_one.append(x)

	# CREATE DATABASE SQLITE WEEK 1
	con = sqlite3.connect("_internal/database/time_sheet.db")
	c = con.cursor()
	# Create table
	c.execute(
		"""CREATE TABLE IF NOT EXISTS rosterAttWeekOne (
			name TEXT,
			badge TEXT,
			thur TEXT,
			fri TEXT,
			sat TEXT,
			sun TEXT,
			mon TEXT,
			tue TEXT,
			wed TEXT
			)"""
	)

	# Add week one data to table
	week1 = """INSERT INTO rosterAttWeekOne (
			name,
			badge,
			thur,
			fri,
			sat,
			sun,
			mon,
			tue,
			wed
			)
			VALUES (?, ?, ?, ?, ? ,? ,?, ?, ?)"""

	c.execute(week1, ("Date", "999", 0, 0, 0, 0, 0, 0, 0))

	for week in week_one:
		x = (week[0], 0, week[1], week[2], week[3], week[4], week[5], week[6], week[7])
		c.execute(week1, (x))

		con.commit()

	# UPDATE TABLE WITH DATES FROM ROSTER
	query = """UPDATE rosterAttWeekOne SET
			thur = ?,
			fri = ?,
			sat = ?,
			sun = ?,
			mon = ?,
			tue = ?,
			wed = ?
			WHERE badge = ?
			"""
	thursday = weekone_dates["Thursday"]
	friday = weekone_dates["Friday"]
	saturday = weekone_dates["Saturday"]
	sunday = weekone_dates["Sunday"]
	monday = weekone_dates["Monday"]
	tuesday = weekone_dates["Tuesday"]
	wednesday = weekone_dates["Wednesday"]

	c.execute(
		query, (thursday, friday, saturday, sunday, monday, tuesday, wednesday, 999)
	)
	con.commit()

	# GET ALL INFO FROM DATABASE FOR PROCESSING
	c.execute("SELECT name FROM rosterAttWeekOne")
	name_records = c.fetchall()

	week_one_data_db = []

	for record in name_records:
		# Grab data from database using name of person
		c.execute("SELECT * FROM rosterAttWeekOne where name=?", (record[0],))
		r = c.fetchall()
		if record[0] != 'Date':
			week_one_data_db.append(r)

	con.close()
	
	return week_one_data_db, weekone_dates

# ==============================================================================
# ATTENDENTS WEEK 2
# ==============================================================================

def attendant_two(selected):
	# Get Dates
	file = '../Attendant_Carwash_Roster.xlsx'
	file_sheets = pd.ExcelFile(file).sheet_names

	# Get Columns
	columns = ['idx','ATTENDANTS', 'THURS', 'FRI', 'SAT', 'SUN', 'MON', 'TUE', 'WED']
	
	# Get Selected sheet
	index = file_sheets.index(selected)
	sheet = file_sheets[index]

	# Get Times
	data = pd.read_excel(file, sheet_name=sheet, index_col=0, header=1, usecols=columns)
	data = data.fillna(0)

	# Get dates
	data_date = pd.read_excel(file, sheet_name=sheet, header=None, usecols='C:I').loc[28]

	weektwo_dates = {}
	for x in data_date:
		weektwo_dates[x.strftime("%A")] = x.date().strftime("%d/%m/%y")

	# Get week one from excel roster sheet
	week_two_data = data.loc[30:44]
	week_two = []
	for x in week_two_data.to_numpy().tolist():
		if str(x[0]) != 'nan':
			if x[0] != 0:
				week_two.append(x)

	# CREATE DATABASE SQLITE WEEK 2
	con = sqlite3.connect("_internal/database/time_sheet.db")
	c = con.cursor()
	# Create table
	c.execute(
		"""CREATE TABLE IF NOT EXISTS rosterAttWeekTwo (
			name TEXT,
			badge TEXT,
			thur TEXT,
			fri TEXT,
			sat TEXT,
			sun TEXT,
			mon TEXT,
			tue TEXT,
			wed TEXT
			)"""
	)

	# Add week one data to table
	week2 = """INSERT INTO rosterAttWeekTwo (
			name,
			badge,
			thur,
			fri,
			sat,
			sun,
			mon,
			tue,
			wed
			)
			VALUES (?, ?, ?, ?, ? ,? ,?, ?, ?)"""

	c.execute(week2, ("Date", "999", 0, 0, 0, 0, 0, 0, 0))

	for week in week_two:
		x = (week[0], 0, week[1], week[2], week[3], week[4], week[5], week[6], week[7])
		c.execute(week2, (x))

		con.commit()

	# UPDATE TABLE WITH DATES FROM ROSTER
	con = sqlite3.connect("_internal/database/time_sheet.db")
	c = con.cursor()

	query = """UPDATE rosterAttWeekTwo SET
			thur = ?,
			fri = ?,
			sat = ?,
			sun = ?,
			mon = ?,
			tue = ?,
			wed = ?
			WHERE badge = ?
			"""
	thursday = weektwo_dates["Thursday"]
	friday = weektwo_dates["Friday"]
	saturday = weektwo_dates["Saturday"]
	sunday = weektwo_dates["Sunday"]
	monday = weektwo_dates["Monday"]
	tuesday = weektwo_dates["Tuesday"]
	wednesday = weektwo_dates["Wednesday"]

	c.execute(
		query, (thursday, friday, saturday, sunday, monday, tuesday, wednesday, 999)
	)
	con.commit()

	# GET ALL INFO FROM DATABASE FOR PROCESSING
	c.execute("SELECT name FROM rosterAttWeekTwo")
	name_records = c.fetchall()

	week_two_data_db = []

	for record in name_records:
		# Grab data from database using name of person
		c.execute("SELECT * FROM rosterAttWeekTwo where name=?", (record[0],))
		r = c.fetchall()
		if record[0] != 'Date':
			week_two_data_db.append(r)

	con.close()

	return week_two_data_db, weektwo_dates