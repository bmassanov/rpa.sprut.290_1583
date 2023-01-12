from datetime import datetime
from os import environ
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv

load_dotenv()
app_base = environ.get('app_base')
app_login = environ.get('app_login')
app_pass = environ.get('app_pass')
owa_login = environ.get('owa_login')
owa_pass = environ.get('owa_pass')

if datetime.now().month != 1:
    start_date = (datetime.now() - relativedelta(months=1)).replace(day=30).strftime('%d.%m.%Y')
else:
    start_date = datetime.now().replace(day=1).strftime('%d.%m.%Y')
end_date = datetime.now().replace(day=30).strftime('%d.%m.%Y')
report_month = datetime.now().month
