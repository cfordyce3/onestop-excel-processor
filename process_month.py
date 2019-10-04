import os
from process_week import process_week
from error_check import error_check, error_correct

FILE_DIR = os.path.dirname(os.path.realpath(__file__))

def process_month (year='', month=''):
    if (month==''): year=input('What year? ')
    if (month==''): month=input('What month? ')

    print()
    print('Processing data for ' + month + ', ' + str(year) + '... ', end='')

    files = []

    for file in os.listdir(FILE_DIR + '\\Statistics\\Master\\' + str(year) + '\\Weekly\\' + month):
        day = file.strip('Week of ' + month + ' '); day = day.rstrip(str(year) + '.xlsx'); day = day.strip(", ")
        files.append([str(day),str(month),str(year)])
        #print(files)

    for file in files:
        process_week(file[0],file[1],file[2])

    print('Done.')
    print()

    # error checking
    print('Error checking processed files.')
    for file in files:
        if (error_check(file[0],file[1],file[2]) == True):
            print(' ~ Errors found in \'' + str(file) + '\' Correcting... ', end = '')
            error_correct(file[0],file[1],file[2])
            print('Done.')

    print()
    print('Data for ' + month + ', ' + str(year) + ' has been completed.')
    print()

#process_month()
