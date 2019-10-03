import os
from process_week import process_week

def process_month (year='', month=''):
    if (month==''): year=input('What year? ')
    if (month==''): month=input('What month? ')

    for file in os.listdir('C:\\Users\\Chas\\Documents\\GitHub\\onestop-excel-processor\\Statistics\\Master\\' + str(year) + '\\Weekly\\' + month):
        day = file.strip('Week of ' + month + ' ')
        day = day.rstrip(str(year) + '.xlsx')
        day = day.strip(", ")
        process_week(day,month,year)

process_month()
