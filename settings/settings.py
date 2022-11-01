from os import environ

from dotenv import load_dotenv

load_dotenv()
app_base = environ.get('app_base')
app_login = environ.get('app_login')
app_pass = environ.get('app_pass')
owa_login = environ.get('owa_login')
owa_pass = environ.get('owa_pass')

start_date = '30.09.2022'
end_date = '30.10.2022'
report_month = 10
