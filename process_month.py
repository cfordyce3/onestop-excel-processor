import os
from process_week import process_week
from resources import get_stats_folder

FILE_DIR = get_stats_folder()

def process_month (month='',year='',show_info=1):
    if (show_info==1):
        print()
        print('Processing data for ' + month + ', ' + str(year) + '... ', end='')

    files = []

    for file in os.listdir(FILE_DIR + '\\Statistics\\Masters\\Weekly\\' + str(year) + '\\' + month):
        day = file.strip('Week of ' + month + ' '); day = day.rstrip(str(year) + '.xlsx'); day = day.strip(", ")
        files.append([str(day),str(month),str(year)])
        #print(files)

    for file in files:
        process_week(file[0],file[1],file[2],show_info)

    if (show_info==1):
        print('Done.')
        print()

    if (show_info==1):
        print()
        print('Data for ' + month + ', ' + str(year) + ' has been completed.')
        print()

#process_month('December','2019', 1)
