import os
from openpyxl import load_workbook

FILE_DIR = os.path.abspath('..')
months = {'January':1,'February':2,'March':3,'April':4,'May':5,'June':6,'July':7,'August':8,'September':9,'October':10,'November':11,'December':12}

def create_list_of_inputs (start_week='',start_month='',start_year='',end_week='',end_month='',end_year='',show_info=False):
    start_month_val = months[start_month]
    end_month_val = months[end_month]
    print (start_month_val)
    print (end_month_val)
    inputs = []

    if (start_year==end_year):
        for month in months:
            if (month.value() >= start_month_val and month.value() <= end_month_val):
                for file in os.listdir(FILE_DIR + '\\Statistics\\Master\\' + str(year) + '\\Weekly\\' + month.key()):
                    print(file)

def process_quarter (season='',start_week='',start_month='',start_year='',end_week='',end_month='',end_year='',show_info=False):
    if (season=='' and start_week=='' and start_month=='' and start_year=='' and end_week=='' and end_month=='' and end_year==''):
        start_month=input('What starting month? '); start_week=str(input('What starting week? ')); start_year=input('What starting year? ')
        end_month=input('What ending month? '); end_week=str(input('What ending week? ')); end_year=input('What ending year? ')
        season=input('What season are we processing? ')

    out_file = FILE_DIR + '\\Statistics\\Master\\' + end_year + '\\Quarterly\\' + season + ' ' + end_year + '.xlsx'
    outwb = load_workbook(filename = out_file)

    input_files = create_list_of_inputs(start_week,start_month,start_year,end_week,end_month,end_year,show_info)

    outwb.save(out_file)


create_list_of_inputs('20','June','2019','9','September','2019',True)
#process_quarter('Summer','20','June','2019','9','September','2019',True)