import pandas as pd
import re

def bakery_one(selected):
    bakery_week_one_list = []
    
    # Read in excel file
    file = '../CASHIERS_ROSTER.xls'
    file_sheets = pd.ExcelFile(file).sheet_names
    # Get Columns
    columns = ['idx', 'CASHIERS', 'THU', 'FRI', 'SAT', 'SUN', 'MON', 'TUE', 'WED']
    # Get last sheet
    sheet = file_sheets[selected]
    # Read in file last sheet
    data = pd.read_excel(file, sheet_name=sheet, index_col=0, header=4, usecols=columns)

    # Get week one from excel sheet
    week_one_data = data.loc[13:14]
    week_one = week_one_data.to_numpy()
    
    # FUNCTIONS
    # Find numbers before dash and after dash
    def first(weekday):
        first = float(re.findall('[0-9.]+(?=.*\-)', weekday)[0])
        return first

    def second(weekday):
        second = float(re.findall('\-(.*)', weekday)[0])
        return second

    # WORKINGS
    for week in week_one:
        name = week[0]
        thur = week[1]
        fri = week[2]
        sat = week[3]
        sun = week[4]
        mon = week[5]
        tue = week[6]
        wed = week[7]

        if thur == 'AF':
            thur = 0
        elif first(thur) == 6.3:
            thur = second(thur) - 6.5    
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
        elif first(fri) == 6.3:
            fri = second(fri) - 6.5  
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
        elif first(sat) == 6.3:
            sat = second(sat) - 6.5
            
        elif second(sat) == 6.3:
            sat = (24 - first(sat)) + 6.5
        elif first(sat) >= 18:
            sat = (24 - first(sat)) + second(sat)
        elif first(sat) - second(sat) < 0:
            sat = (first(sat) - second(sat)) * -1    
        else:
            sat = first(sat) - second(sat)

        if sun == 'AF':
            sun = 0
        elif first(sun) == 6.3:
            sun = second(sun) - 6.5  
        elif second(sun) == 6.3:
            sun = (24 - first(sun)) + 6.5
        elif first(sun) >= 18:
            sun = (24 - first(sun)) + second(sun)
        elif first(sun) - second(sun) < 0:
            sun = (first(sun) - second(sun)) * -1
        else:
            sun = first(sun) - second(sun)

        if mon == 'AF':
            mon = 0
        elif first(mon) == 6.3:
            mon = second(mon) - 6.5  
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
        elif first(tue) == 6.3:
            tue = second(tue) - 6.5  
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
        elif first(wed) == 6.3:
            wed = second(wed) - 6.5  
        elif second(wed) == 6.3:
            wed = (24 - first(wed)) + 6.5
        elif first(wed) >= 18:
            wed = (24 - first(wed)) + second(wed)
        elif first(wed) - second(wed) < 0:
            wed = (first(wed) - second(wed)) * -1
        else:
            wed = first(wed) - second(wed)

        total_hours_week_one = thur + fri + sat + mon + tue + wed
        bakery_week_one_list.append([name , total_hours_week_one , sun]) 
    return bakery_week_one_list

def bakery_two(selected):
    bakery_week_two_list = []

    file = '../CASHIERS_ROSTER.xls'
    file_sheets = pd.ExcelFile(file).sheet_names
    # Get Columns
    columns = ['idx', 'CASHIERS', 'THU', 'FRI', 'SAT', 'SUN', 'MON', 'TUE', 'WED']
    # Get last sheet
    sheet = file_sheets[selected]
    # Read in file last sheet
    data = pd.read_excel(file, sheet_name=sheet, index_col=0, header=4, usecols=columns)

    # Get week two from excel sheet
    week_two_data = data.loc[15:16]
    week_two = week_two_data.to_numpy()
    
    # FUNCTIONS
    # Find numbers before dash and after dash
    def first(weekday):
        first = float(re.findall('[0-9.]+(?=.*\-)', weekday)[0])
        return first

    def second(weekday):
        second = float(re.findall('\-(.*)', weekday)[0])
        return second
    
    # WORKINGS
    for week in week_two:
        name = week[0]
        thur = week[1]
        fri = week[2]
        sat = week[3]
        sun = week[4]
        mon = week[5]
        tue = week[6]
        wed = week[7]

        if thur == 'AF':
            thur = 0
        elif first(thur) == 6.3:
            thur = second(thur) - 6.5    
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
        elif first(fri) == 6.3:
            fri = second(fri) - 6.5  
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
        elif first(sat) == 6.3:
            sat = second(sat) - 6.5  
        elif second(sat) == 6.3:
            sat = (24 - first(sat)) + 6.5
        elif first(sat) >= 18:
            sat = (24 - first(sat)) + second(sat)
        elif first(sat) - second(sat) < 0:
            sat = (first(sat) - second(sat)) * -1
        else:
            sat = first(sat) - second(sat)

        if sun == 'AF':
            sun = 0
        elif first(sun) == 6.3:
            sun = second(sun) - 6.5  
        elif second(sun) == 6.3:
            sun = (24 - first(sun)) + 6.5
        elif first(sun) >= 18:
            sun = (24 - first(sun)) + second(sun)
        elif first(sun) - second(sun) < 0:
            sun = (first(sun) - second(sun)) * -1
        else:
            sun = first(sun) - second(sun)

        if mon == 'AF':
            mon = 0
        elif first(mon) == 6.3:
            mon = second(mon) - 6.5  
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
        elif first(tue) == 6.3:
            tue = second(tue) - 6.5  
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
        elif first(wed) == 6.3:
            wed = second(wed) - 6.5  
        elif second(wed) == 6.3:
            wed = (24 - first(wed)) + 6.5
        elif first(wed) >= 18:
            wed = (24 - first(wed)) + second(wed)
        elif first(wed) - second(wed) < 0:
            wed = (first(wed) - second(wed)) * -1
        else:
            wed = first(wed) - second(wed)

        total_hours_week_two = thur + fri + sat + mon + tue + wed
        bakery_week_two_list.append([name , total_hours_week_two , sun])
    return bakery_week_two_list



