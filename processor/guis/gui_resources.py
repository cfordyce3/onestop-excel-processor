import datetime

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