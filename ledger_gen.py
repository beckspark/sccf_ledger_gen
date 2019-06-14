from __future__ import print_function
import win32com.client as win32
import shutil
from openpyxl import load_workbook
import pandas as pd
import pyperclip
import datetime
from datetime import datetime, timedelta
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import subprocess
from time import sleep
import pywinauto
from pywinauto import *

# This script will generate the ledger for a mobile dental clinic at SCCF

# First, let's take today's date and the next 3 days after, and ask the person which date they want to create the ledger for

creds = None
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.

if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server()
    # Save the credentials for the next run

    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

nday = int(input('How many days ahead do you want to find the clinic?: '))
service = build('calendar', 'v3', credentials=creds)
today = datetime.today()
limit = today + timedelta(days=nday)
eventmin = today.isoformat() + 'Z'
eventmax = limit.isoformat() + 'Z'


# Now that we have our dates establish, we'll establish a connection with the MDC Calendar in Google Calendar, get the events listed, and then ask the user to choose which ledger they want to generate.

events_result = service.events().list(calendarId='5eg17vp0ot0b3mlfh2v01tjra0@group.calendar.google.com', timeMin=eventmin, timeMax=eventmax, singleEvents=True, orderBy='startTime').execute()
events = events_result.get('items', [])

if not events:
    print('No upcoming events found.')
    SystemExit()
for event in events:
    start = event['start'].get('dateTime', event['start'].get('date'))
    clinic = start," ", event['summary']
    clinic = ''.join(clinic)
    clinic_d = datetime.strptime(start, "%Y-%m-%d")
    clinic_d = clinic_d.strftime('%m/%d/%Y')
    cont = 'nah'
    while cont != 'YES' and cont != 'NO':
        cont = input('Do you wish to generate a ledger for ' + clinic + '?: ')
    if cont == 'YES':
        break
if cont == 'NO':
    print('No ledger to process!')
    SystemExit()

# Next we'll take the date of the chosen clinic, have datetime reorganize it so DentiMax can read it, and pull the appointment from DentiMax. Once we have names, we'll cross-reference them with the hourly_rate csv, and add everyone's name and hourly rate to the ledger template, saving it with the name and date of the clinic. We can also add a feature to send Melanie an e-mail automatically asking her to add the balances.
dMaxFile = os.path.join('cpub-dentimax-APP7-CmsRdsh.rdp')
subprocess.call(['taskkill','/f','/im', 'mstsc.exe', '/t'])
dMax = subprocess.Popen(['mstsc',dMaxFile], shell=True)
cont = 'no'
while cont != 'YES':
    cont = input('DentiMax will now start in a separate window. Please log in and navigate to the home page. Enter \'YES\' to continue: ')
app = application.Application()
h = pywinauto.findwindows.find_windows(title_re=".*GATEWAY1",class_name="RAIL_WINDOW")
if len(h) > 1:
    app.connect(handle=h[1])
elif len(h) <= 1:
    app.connect(handle=h[0])
dmWin = app.top_window()
dmWin.type_keys('{F4}') # Returns to home screen in DentiMax just in case
sleep(3)
dmWin.type_keys('%')
sleep(.1)
dmWin.type_keys('{RIGHT}')
sleep(.1)
dmWin.type_keys('{RIGHT}')
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{ENTER}')
sleep(3)
dmWin.type_keys('{TAB}')
sleep(.1)
dmWin.type_keys('{TAB}')
sleep(.1)
dmWin.type_keys('{TAB}')
sleep(.1)
dmWin.type_keys('{TAB}')
sleep(.1)
dmWin.type_keys('{TAB}')
sleep(.1)
dmWin.type_keys('{TAB}')
sleep(.1)
pyperclip.copy(clinic_d)
dmWin.type_keys('^v')
sleep(.1)
dmWin.type_keys('{TAB}')
sleep(1)
dmWin.click(button="right",coords=(300,320))
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{DOWN}')
sleep(.1)
dmWin.type_keys('{ENTER}')
sleep(2)

c = pywinauto.findwindows.find_windows(title_re=".*Confirm",class_name="RAIL_WINDOW")
if len(c) > 1:
    app.connect(handle=c[1])
elif len(h) <= 1:
    app.connect(handle=c[0])
dmcWin = app.top_window()
dmcWin.type_keys('{ENTER}')
sleep(2)

e = pywinauto.findwindows.find_windows(title_re=".*dentimaxexport",class_name="RAIL_WINDOW")
dmeWin = app.top_window()
dmeWin.type_keys('^a')
sleep(.1)
dmeWin.type_keys('^c')
sleep(1)

# And now pandas! Pandas will read the clipboard as a csv, add a new column with their full name, then save that as its own dataframe, filtering out OPT4s and any Hygeine folks.
db = pd.read_clipboard(sep=',')
db['fullname'] = db['First Name'] + ' ' + db['Last Name']
dbn = db[db.Resource == 'OPT1'].append(db[db.Resource == 'OPT2']).append(db[db.Resource == 'OPT3']).fullname
dbn = dbn.to_list()
del dbn[0]
subprocess.call(['taskkill','/f','/im', 'mstsc.exe', '/t'])

# Next up, let's take the csv hourly rate file and add the hourly rates
dbh = []
hdb = pd.read_csv('hourly_rates.csv')
for n in dbn:
    try:
        i = hdb[hdb.fullname == n].hourly.item()
        print(i)
    except:
        i = int(input('What is the hourly rate for ' + n + '?: ').replace("$",""))
        hdb = hdb.append(pd.DataFrame([[n,i]], columns=hdb.columns))
        hdb.to_csv('hourly_rates.csv', index=False)
        hdb = pd.read_csv('hourly_rates.csv')
        pass
    dbh.append(i)

tledger = os.path.join(os.getcwd(),"ledger_template.xlsx")
sledger = os.path.join(os.getcwd(), clinic + '.xlsx')
shutil.copyfile(tledger,sledger)

xl = win32.gencache.EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(sledger)
ws = wb.Worksheets(1)
row = 5

for i in range(len(dbn)):
    i = i
    ws.Range('A' + str(row)).Value = str(dbn[i])
    i = i+1
    row = row+1

row = 5

for i in range(len(dbh)):
    i = i
    ws.Range('C' + str(row)).Value = str(dbh[i])
    i = i+1
    row = row+1

ws.Columns.AutoFit()
wb.Save()
xl.Quit()

# To finish, let's go ahead and move the ledger to the correct directory!
dbox = "\\\\10.1.10.201\\Dropbox\\Services\\Approval Letters for Charity Care\\"
clinic_d = datetime.strptime(start, "%Y-%m-%d")
year = clinic_d.strftime('%Y')
month = clinic_d.strftime('%B')
db_path = os.path.join(dbox,year,"Daily Accounting Ledgers",month)
if not os.path.exists(os.path.join(db_path)):
    os.makedirs(db_path)

wledger = os.path.join(db_path,(clinic + '.xlsx'))
shutil.copy(sledger,wledger)

cont = 'no'
while cont != 'y':
    cont = input('Ledger is processed! Enter \'y\' to exit: ')
SystemExit()
