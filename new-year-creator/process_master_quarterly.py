import os
from shutil import copyfile
from datetime import date, timedelta
from openpyxl import Workbook, load_workbook
from resources import get_stats_folder

path = get_stats_folder()

print("What year? ")
year = 2021#int(input())

def get_mondays (year):
    sdate = date(year, 1, 1)
    edate = date(year, 12, 31)
    deltadate = edate - sdate

    mondays = []

    for i in range(deltadate.days + 1):
        day = sdate + timedelta(days=i)
        if (day.weekday() == 0):
            mondays.append(str(day.month) + "." + str(day.day))

    return mondays

mondays = get_mondays(year)

# make quarterly directory
#'\\Statistics\\Masters\\Quarterly\\'
working_path = path + '\\new-year-creator\\' + str(year)
os.mkdir(working_path)

copyfile((path + '\\Statistics\\Masters\\Quarterly\\Template\\Quarterly Template.xlsx'), working_path)
