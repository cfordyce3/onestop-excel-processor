# import required libraries
import os
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter

# set working directory to the directory above where this file is located
FILE_DIR = os.path.abspath('..')

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

# correct the totals at the end
def correct_totals (wb):
    # correct the daily sheets
    for sheet_name in sheet_day:
        ws = wb[sheet_name]
        for x in range(2,15):
            for y in range(51,55):
                ws.cell(row=y,column=x).value = '=SUM(' + ws.cell(row=y-48,column=x).coordinate + '+' + ws.cell(row=y-39,column=x).coordinate + '+' + ws.cell(row=y-30,column=x).coordinate + '+' + ws.cell(row=y-21,column=x).coordinate + '+' + ws.cell(row=y-12,column=x).coordinate + ')'

    # correct the weekly stats section up to the final total section
    multiplier = 0
    for sheet_name in sheet_day:
        ws = wb['Weekly Stats']

        for x in range(2,15):
            y_source = 51
            for y in range(3+(multiplier*7),7+(multiplier*7)):
                if (y != (7+(multiplier*7)-1)):
                    ws.cell(row=y,column=x).value = str('=' + sheet_name + '!' + get_column_letter(x) + str(y_source))
                    y_source += 1
                else:
                    ws.cell(row=y,column=x).value = str('=SUM(' + get_column_letter(x) + str(y-3) + ':' + get_column_letter(x) + str(y-1) + ')')
        multiplier += 1

        # set up final total section at the bottom of the sheet
        for x in range(2,15):
            for y in range(38,42):
                if (y != 41):
                    ws.cell(row=y,column=x).value = str('=SUM(' + get_column_letter(x) + str(y-35) + ',' + get_column_letter(x) + str(y-28) + ',' + get_column_letter(x) + str(y-21) + ',' + get_column_letter(x) + str(y-14) + ',' + get_column_letter(x) + str(y-7) + ')')
                else:
                    ws.cell(row=y,column=x).value = str('=SUM(' + get_column_letter(x) + str(y-3) + ':' + get_column_letter(x) + str(y-1) + ')')


#wb = load_workbook(filename='C:\\Users\\Chas\\Documents\\GitHub\\onestop-excel-processor\\Statistics\\Master\\2019\\Weekly\\August\\Week of August 26, 2019.xlsx')
#correct_totals(wb)
#wb.save('C:\\Users\\Chas\\Documents\\GitHub\\onestop-excel-processor\\Statistics\\Master\\2019\\Weekly\\August\\Week of August 26, 2019.xlsx')



def process_week (week='',month='',year='',show_info=False):

    if (week=='' and month=='' and year==''):
        month=input('What month? '); week=str(input('What week? ')); year=input('What year? ')

    if (show_info == True):
        print()

    out_file = FILE_DIR + '\\Statistics\\Master\\' + year + '\\Weekly\\' + month + '\\Week of ' + month + ' ' + week + ', ' + year + '.xlsx'
    outwb = load_workbook(filename = out_file)

    if (show_info == True):
        print('[' + week + ' ' + month + ', ' + year + ']')

    for os_num in range(0,5):
        in_file = FILE_DIR + '\\Statistics\\'
        if (os_num == 0):
            if (show_info == True):
                print('Processing Low Desk... ', end='')
            in_file += 'Low Desk\\' + year + '\\' + month + '\\LD Week of ' + month + ' ' + week + ', ' + year + '.xlsx'
        else:
            if (show_info == True):
                print('Processing OS' + str(os_num) + '... ', end='')
            in_file += 'OS' + str(os_num) + '\\' + year + '\\' + month + '\\OS' + str(os_num) + ' Week of ' + month + ' ' + week + ', ' + year + '.xlsx'

        try:
            f = open(in_file)
            f.close()
        except FileNotFoundError:
            if (show_info == True):
                print('FILE NOT FOUND. DATA NOT COMPILED!')
            continue

        inwb = load_workbook(filename = in_file, data_only = True)
        for sheet_name in sheet_day:
            inws = inwb[sheet_name]
            outws = outwb[sheet_name]
            row_thru(os_num,inws,outws)

        if (show_info == True):
            print('Done.')


    correct_totals(outwb)

    # determine save file and saves file
    save_file = FILE_DIR + '\\Statistics\\Master\\' + year + '\\Weekly\\' + month + '\\Week of ' + month + ' ' + week + ', ' + year + '.xlsx'
    outwb.save(save_file)



#process_week('26','August','2019')
