import os
import copy
from openpyxl import load_workbook
from openpyxl.chart import PieChart3D, Reference

FILE_DIR = os.path.abspath('..')
months = {'January':1,'February':2,'March':3,'April':4,'May':5,'June':6,'July':7,'August':8,'September':9,'October':10,'November':11,'December':12}

def create_dict_of_inputs (start_week='',start_month='',start_year='',end_week='',end_month='',end_year='',show_info=False):
    start_month_val = months[start_month]
    end_month_val = months[end_month]

    inputs = {}

    if (start_year==end_year):
        year_list = {}
        month_list = {}
        for month_name,month_num in months.items():
            day_list = []
            if (month_num >= start_month_val and month_num <= end_month_val):
                for file in os.listdir(FILE_DIR + '\\Statistics\\Master\\' + str(start_year) + '\\Weekly\\' + month_name):
                    day = file.strip('Week of ' + month_name + ' '); day = day.rstrip(str(start_year) + '.xlsx'); day = day.strip(", ")
                    if (month_name != start_month and month_name != end_month):
                        day_list.append(file)
                    elif (month_name == end_month and int(day) <= int(end_week)):
                        day_list.append(file)
                    elif (month_name == start_month and int(day) >= int(start_week)):
                        day_list.append(file)
                if (len(day_list) != 0):
                    month_list[month_name] = day_list
            if (len(month_list) != 0):
                year_list[start_year] = month_list
        inputs = year_list

    if (show_info == True):
        print(inputs)

    return inputs

def find_worksheet (wb,ws_name):
    for sheet in wb:
        if (sheet.title == ws_name):
            return sheet
        if (sheet.title.find('2015') != -1 or sheet.title.find('2014') != -1):
            return sheet

    return False

def create_new_worksheet (wb,ws_name):
    wb.add_sheet(copy.deepcopy(ws),ido+1)
    #ws = wb.copy_worksheet(wb['111111'])
    ws.title = ws_name

    '''pie_chart = PieChart3D()
    labels = Reference(ws,min_col=2,max_col=5,min_row=9,max_row=9)

    data = [['Bursar',        '=SUM(B6:D6)'],
            ['Financial Aid', '=SUM(E6:G6)'],
            ['Registrar',     '=SUM(H6:J6)'],
            ['Other',         '=SUM(K6:M6)']]

    r = 20; c = 2
    for row in data:
        ws.cell(row=r,column=c,value=row[0])
        c += 1
        ws.cell(row=r,column=c,value=row[1])
        c -= 1
        r += 1

    data = Reference(ws,min_col=2,max_col=3,min_row=20,max_row=23)

    for row in ws['B20':'C23']:
        for cell in row:
            cell.value = None

    data = Reference(ws,min_col=2,max_col=2,)
    pie_chart.add_data(data,titles_from_data=True)
    pie_chart.set_categories(labels)
    pie_chart.legend.position = 'b'
    pie_chart.title = ''

    ws.add_chart(pie_chart,'G9')'''


def row_thru (outwb,year,month,filename,worksheet_count,show_info=False):
    in_file = FILE_DIR + '\\Statistics\\Master\\' + year + '\\Weekly\\' + month + '\\' + filename
    inwb = load_workbook(filename=in_file)
    inws = inwb['Weekly Stats']

    ws_name_from_filename = filename.rstrip('.xlsx')
    outws = find_worksheet(outwb,ws_name_from_filename)

    if (outws == False):
        create_new_worksheet(outwb,ws_name_from_filename)





def process_quarter (season='',start_week='',start_month='',start_year='',end_week='',end_month='',end_year='',show_info=False):
    if (season=='' and start_week=='' and start_month=='' and start_year=='' and end_week=='' and end_month=='' and end_year==''):
        start_month=input('What starting month? '); start_week=str(input('What starting week? ')); start_year=input('What starting year? ')
        end_month=input('What ending month? '); end_week=str(input('What ending week? ')); end_year=input('What ending year? ')
        season=input('What season are we processing? ')

    out_file = FILE_DIR + '\\Statistics\\Master\\' + end_year + '\\Quarterly\\' + season + ' ' + end_year + '.xlsx'
    outwb = load_workbook(filename = out_file)

    input_files = create_dict_of_inputs(start_week,start_month,start_year,end_week,end_month,end_year,show_info)

    for year,month_list in input_files.items():
        for month,actual_files in month_list.items():
            for file in actual_files:
                row_thru(outwb,year,month,file,show_info)


    outwb.save(out_file)


#create_dict_of_inputs('20','June','2019','9','September','2019',True)
process_quarter('Spring','20','June','2019','9','September','2019',False)
