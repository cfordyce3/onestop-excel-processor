import os
from openpyxl import load_workbook

FILE_DIR = os.path.abspath('..')
months = {'January':1,'February':2,'March':3,'April':4,'May':5,'June':6,'July':7,'August':8,'September':9,'October':10,'November':11,'December':12}

def create_dict_of_inputs (start_week='',start_month='',start_year='',end_week='',end_month='',end_year='',show_info=False):
    start_month_val = months[start_month]
    end_month_val = months[end_month]

    inputs = {}

    if (start_year==end_year):
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
                    inputs[month_name] = day_list

    if (show_info == True):
        print(inputs)

    return inputs

def process_quarter (season='',start_week='',start_month='',start_year='',end_week='',end_month='',end_year='',show_info=False):
    if (season=='' and start_week=='' and start_month=='' and start_year=='' and end_week=='' and end_month=='' and end_year==''):
        start_month=input('What starting month? '); start_week=str(input('What starting week? ')); start_year=input('What starting year? ')
        end_month=input('What ending month? '); end_week=str(input('What ending week? ')); end_year=input('What ending year? ')
        season=input('What season are we processing? ')

    out_file = FILE_DIR + '\\Statistics\\Master\\' + end_year + '\\Quarterly\\' + season + ' ' + end_year + '.xlsx'
    outwb = load_workbook(filename = out_file)

    input_files = create_dict_of_inputs(start_week,start_month,start_year,end_week,end_month,end_year,show_info)

    for month,actual_files in input_files.items():
        pass


    outwb.save(out_file)


create_dict_of_inputs('20','June','2019','9','September','2019',True)
#process_quarter('Summer','20','June','2019','9','September','2019',True)
