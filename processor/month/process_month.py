import os
from processor.week.process_week import process_week

import processor

FILE_DIR = os.path.join(os.path.dirname(processor.__file__),'..') #os.path.abspath('..')

def process_month (month='',year='',show_info=0):
    if (show_info==1):
        print()
        print('Processing data for ' + month + ', ' + str(year) + '... ', end='')

    files = []

    for file in os.listdir(FILE_DIR + '\\Statistics\\Master\\' + str(year) + '\\Weekly\\' + month):
        day = file.strip('Week of ' + month + ' '); day = day.rstrip(str(year) + '.xlsx'); day = day.strip(", ")
        files.append([str(day),str(month),str(year)])
        #print(files)

    for file in files:
        process_week(file[0],file[1],file[2])

    if (show_info==1):
        print('Done.')
        print()

    if (show_info==1):
        print()
        print('Data for ' + month + ', ' + str(year) + ' has been completed.')
        print()


#process_month('August','2019',0)
