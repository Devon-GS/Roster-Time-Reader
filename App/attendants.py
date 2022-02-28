import pandas as pd
import re

# Notes
# before dash [0-9]+(?=.*\-)
# after dash \-(.*)

def attendant_one(selected):
    week_one_list = []

    # Read file and get sheets
    file = '../Attendant_Carwash_Roster.xls'
    file_sheets = pd.ExcelFile(file).sheet_names
    # Get Columns
    columns = ['idx','ATTENDANTS', 'THURS', 'FRI', 'SAT', 'SUN', 'MON', 'TUE', 'WED']
    # Get last sheet
    sheet = file_sheets[selected]
    # Read in file last sheet
    data = pd.read_excel(file, sheet_name=sheet, index_col=0, header=1, usecols=columns)

    # Get week one from excel sheet
    week_one_data = data.loc[0:14]
    week_one = []
    for x in week_one_data.to_numpy():
        if str(x[0]) != 'nan':
            week_one.append(x)
    
    # FUNCTIONS
    # Find numbers before dash and after dash
    def first(weekday):
        first = float(re.findall('[0-9]+(?=.*\-)', weekday)[0])
        return first

    def second(weekday):
        second = float(re.findall('\-(.*)', weekday)[0])
        return second

    # Workings
    for week in week_one:
        name = week[0]
        thur = week[1]
        fri = week[2]
        sat = week[3]
        sun = week[4]
        mon = week[5]
        tue = week[6]
        wed = week[7]
        sune = 0.0
        mone = 0.0

        if thur == 'AF':
            thur = 0
        elif second(thur) == 6.3:
            thur = (24 - first(thur)) + 6.5
        elif first(thur) >= 18:
            thur = (24 - first(thur)) + second(thur)
        elif first(thur) - second(thur) < 0:
            thur = (first(thur) - second(thur)) * -1
        else:
            thur = first(thur) - second(thur)
        
        if fri == 'AF':
            fri = 0
        elif second(fri) == 6.3:
            fri = (24 - first(fri)) + 6.5
        elif first(fri) >= 18:
            fri = (24 - first(fri)) + second(fri)
        elif first(fri) - second(fri) < 0:
            fri = (first(fri) - second(fri)) * -1
        else:
            fri = first(fri) - second(fri)

        if sat == 'AF':
            sat = 0
        elif sat == '18-6' or sat == '18-7':
            sune += second(sat)
            sat = (24 - first(sat))
        elif sat == "18-6.30":
            sune += 6.5
            sat = (24 - first(sat))
        elif sat == '6.30-18':
            sat = second(sat) - 6.5
        else:
            sat = (first(sat) - second(sat)) * -1

        if sun == 'AF':
            sun = 0
        elif sun == '18-6' or sun == '18-7':
            mone += second(sun)
            sun = (24 - first(sun))
        elif sun == '18-6.30':
            mone += 6.5
            sun = (24 - first(sun))
        elif sun == '6.30-18':
            sun = second(sun) - 6.5
        else:
            sun = (first(sun) - second(sun)) * -1

        if mon == 'AF':
            mon = 0
        elif second(mon) == 6.3:
            mon = (24 - first(mon)) + 6.5
        elif first(mon) >= 18:
            mon = (24 - first(mon)) + second(mon)
        elif first(mon) - second(mon) < 0:
            mon = (first(mon) - second(mon)) * -1
        else:
            mon = first(mon) - second(mon)

        if tue == 'AF':
            tue = 0
        elif second(tue) == 6.3:
            tue = (24 - first(tue)) + 6.5
        elif first(tue) >= 18:
            tue = (24 - first(tue)) + second(tue)
        elif first(tue) - second(tue) < 0:
            tue = (first(tue) - second(tue)) * -1
        else:
            tue = first(tue) - second(tue)

        if wed == 'AF':
            wed = 0
        elif second(wed) == 6.3:
            wed = (24 - first(wed)) + 6.5
        elif first(wed) >= 18:
            wed = (24 - first(wed)) + second(wed)
        elif first(wed) - second(wed) < 0:
            wed = (first(wed) - second(wed)) * -1
        else:
            wed = first(wed) - second(wed)

        total_hours_week_one = thur + fri + sat + mon + mone + tue + wed
        total_sun_week_one = sun + sune
        week_one_list.append([name , total_hours_week_one , total_sun_week_one])
    return week_one_list

def attendant_two(selected):
    week_two_list = []

    # Read file and get sheets
    file = '../Attendant_Carwash_Roster.xls'
    file_sheets = pd.ExcelFile(file).sheet_names
    # Get Columns
    columns = ['idx','ATTENDANTS', 'THURS', 'FRI', 'SAT', 'SUN', 'MON', 'TUE', 'WED']
    # Get last sheet
    sheet = file_sheets[selected]
    # Read in file last sheet
    data = pd.read_excel(file, sheet_name=sheet, index_col=0, header=1, usecols=columns)

    # Get week two from excel sheet
    week_two_data = data.loc[30:44]
    
    week_two = []
    for x in week_two_data.to_numpy():
        if str(x[0]) != 'nan':
            week_two.append(x)

    # FUNCTIONS
    # Find numbers before dash and after dash
    def first(weekday):
        first = float(re.findall('[0-9]+(?=.*\-)', weekday)[0])
        return first

    def second(weekday):
        second = float(re.findall('\-(.*)', weekday)[0])
        return second

    for week in week_two:
        name = week[0]
        thur = week[1]
        fri = week[2]
        sat = week[3]
        sun = week[4]
        mon = week[5]
        tue = week[6]
        wed = week[7]
        sune = 0.0
        mone = 0.0

        if thur == 'AF':
            thur = 0
        elif second(thur) == 6.3:
            thur = (24 - first(thur)) + 6.5
        elif first(thur) >= 18:
            thur = (24 - first(thur)) + second(thur)
        elif first(thur) - second(thur) < 0:
            thur = (first(thur) - second(thur)) * -1
        else:
            thur = first(thur) - second(thur)
        
        if fri == 'AF':
            fri = 0
        elif second(fri) == 6.3:
            fri = (24 - first(fri)) + 6.5
        elif first(fri) >= 18:
            fri = (24 - first(fri)) + second(fri)
        elif first(fri) - second(fri) < 0:
            fri = (first(fri) - second(fri)) * -1
        else:
            fri = first(fri) - second(fri)

        if sat == 'AF':
            sat = 0
        elif sat == '18-6' or sat == '18-7':
            sune += second(sat)
            sat = (24 - first(sat))
        elif sat == "18-6.30":
            sune += 6.5
            sat = (24 - first(sat))
        elif sat == '6.30-18':
            sat = second(sat) - 6.5
        else:
            sat = (first(sat) - second(sat)) * -1

        if sun == 'AF':
            sun = 0
        elif sun == '18-6' or sun == '18-7':
            mone += second(sun)
            sun = (24 - first(sun))
        elif sun == '18-6.30':
            mone += 6.5
            sun = (24 - first(sun))
        elif sun == '6.30-18':
            sun = second(sun) - 6.5
        else:
            sun = (first(sun) - second(sun)) * -1

        if mon == 'AF':
            mon = 0
        elif second(mon) == 6.3:
            mon = (24 - first(mon)) + 6.5
        elif first(mon) >= 18:
            mon = (24 - first(mon)) + second(mon)
        elif first(mon) - second(mon) < 0:
            mon = (first(mon) - second(mon)) * -1
        else:
            mon = first(mon) - second(mon)

        if tue == 'AF':
            tue = 0
        elif second(tue) == 6.3:
            tue = (24 - first(tue)) + 6.5
        elif first(tue) >= 18:
            tue = (24 - first(tue)) + second(tue)
        elif first(tue) - second(tue) < 0:
            tue = (first(tue) - second(tue)) * -1
        else:
            tue = first(tue) - second(tue)

        if wed == 'AF':
            wed = 0
        elif second(wed) == 6.3:
            wed = (24 - first(wed)) + 6.5
        elif first(wed) >= 18:
            wed = (24 - first(wed)) + second(wed)
        elif first(wed) - second(wed) < 0:
            wed = (first(wed) - second(wed)) * -1
        else:
            wed = first(wed) - second(wed)

        total_hours_week_two = thur + fri + sat + mon + mone + tue + wed
        total_sun_week_two = sun + sune
        week_two_list.append([name , total_hours_week_two , total_sun_week_two])
    return week_two_list


