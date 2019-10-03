# import required libraries
from openpyxl import load_workbook
import sys

# set variables
sheet_day = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

# define functions
def row_thru (os,inws,outws) :
    sc = 2
    if   (os == 0):
        sr = 3
    elif (os == 1):
        sr = 12
    elif (os == 2):
        sr = 21
    elif (os == 3):
        sr = 30
    elif (os == 4):
        sr = 39

    r = sr
    c = sc
    for row in inws['B82':'N85']:
        for cell in row:
            outws.cell(row = r, column = c, value = cell.value)
            c += 1

        r += 1
        c = sc

# given_dir should include ONLY up until \Statistics\
def process_week (week='',month='',year=''):

    if (week=='' and month=='' and year==''):
        month=input('What month? '); week=str(input('What week? ')); year=input('What year? ')

    print()

    out_file = 'C:\\Users\\fordy\\Desktop\\One Stop GitHub\\onestop-excel-processor\\Statistics\\Master\\' + year + '\\Weekly\\' + month + '\\Week of ' + month + ' ' + week + ', ' + year + '.xlsx'
    outwb = load_workbook(filename = out_file)

    print('[' + week + ' ' + month + ', ' + year + ']')

    for os_num in range(0,5):
        in_file = 'C:\\Users\\fordy\\Desktop\\One Stop GitHub\\onestop-excel-processor\\Statistics\\'
        if (os_num == 0):
            print('Processing Low Desk... ', end='')
            in_file += 'Low Desk\\' + year + '\\' + month + '\\LD Week of ' + month + ' ' + week + ', ' + year + '.xlsx'
        else:
            print('Processing OS' + str(os_num) + '... ', end='')
            in_file += 'OS' + str(os_num) + '\\' + year + '\\' + month + '\\OS' + str(os_num) + ' Week of ' + month + ' ' + week + ', ' + year + '.xlsx'

        try:
            #print(in_file)
            f = open(in_file)
            f.close()
        except FileNotFoundError:
            print('FILE NOT FOUND. DATA NOT COMPILED!')
            continue

        inwb = load_workbook(filename = in_file, data_only = True)
        for sheet_num in sheet_day:
            inws = inwb[sheet_num]
            outws = outwb[sheet_num]
            row_thru(os_num,inws,outws)

        print('Done.')

    # determine save file and saves file
    save_file = 'C:\\Users\\fordy\\Desktop\\One Stop GitHub\\onestop-excel-processor\\Statistics\\Master\\' + year + '\\Weekly\\' + month + '\\COPY Week of ' + month + ' ' + week + ', ' + year + '.xlsx'
    outwb.save(save_file)

#process_week('September','2','2019')
