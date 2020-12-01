import datetime
import os

now = datetime.datetime.now()

YEARS = []
MONTHS = []
DAYS = []

# years
now_year = now.year
for year in range(-5, 6):
    YEARS.append(now_year + year)

# months
now_month = now.month
MONTHS = ['January', 'February', 'March',
          'April', 'May', 'June',
          'July', 'August', 'September',
          'October', 'November', 'December']

# days
now_day = now.day
DAYS = ['1','2','3','4','5','6','7','8','9','10',
        '11','12','13','14','15','16','17','18','19','20',
        '21','22','23','24','25','26','27','28','29','30',
        '31']

WEEKDAYS = {0:'Monday',
            1:'Tuesday',
            2:'Wednesday',
            3:'Thursday',
            4:'Friday',
            5:'Saturday',
            6:'Sunday'}

# find the Statistics folder to start the working
def get_stats_folder ():
    found = False
    count = 0 # just to make sure no infinite loops occur, including a counter; it should not happen regardless
    while (found == False and count < 20):
        if ("Statistics" in os.listdir()):
            found = True
        else:
            os.chdir('..')
            count+=1
    return os.getcwd()
