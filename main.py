from __future__ import print_function
import gspread
from oauth2client.client import SignedJwtAssertionCredentials
import pandas as pd
import json
import sugarcrm

from applicantsModule import applicantsModule

SCOPE = ["https://docs.google.com/spreadsheets/d/1gyYDqjEGkRrTDdJQGTb3sEgHymJXg49m-Yc7lfKKi94/edit#gid=294310776"]
SECRETS_FILE = "api-project-313216-1dd8217a1523.json"
SPREADSHEET = "Новая форма (Ответы)"
# Based on docs here - http://gspread.readthedocs.org/en/latest/oauth2.html
# Load in the secret JSON key (must be a service account)
json_key = json.load(open(SECRETS_FILE))

credentials = SignedJwtAssertionCredentials(json_key['client_email'],
                                            json_key['private_key'], SCOPE)
gc = gspread.service_account()
#gc = gspread.authorize(credentials)
print("The following sheets are available")
for sheet in gc.openall():
    print("{} - {}".format(sheet.title, sheet.id))

workbook = gc.open(SPREADSHEET)

sheet = workbook.sheet1

data = pd.DataFrame(sheet.get_all_records())

column_names = {'Timestamp': 'timestamp'}
#                'What version of python would you like to see used for the examples on the site?': 'version',
#                'How useful is the content on practical business python?': 'useful',
#                'What suggestions do you have for future content?': 'suggestions',
#                'How frequently do you use the following tools? [Python]': 'freq-py',
#                'How frequently do you use the following tools? [SQL]': 'freq-sql',
#                'How frequently do you use the following tools? [R]': 'freq-r',
#                'How frequently do you use the following tools? [Javascript]': 'freq-js',
#                'How frequently do you use the following tools? [VBA]': 'freq-vba',
#                'How frequently do you use the following tools? [Ruby]': 'freq-ruby',
#                'Which OS do you use most frequently?': 'os',
#                'Which python distribution do you primarily use?': 'distro',
#                'How would you like to be notified about new articles on this site?': 'notify'
#                }
#data.rename(columns=column_names, inplace=True)
#data.timestamp = pd.to_datetime(data.timestamp)
print(data.head())

session = sugarcrm.Session("https://crm.uni-dubna.ru/service/v4/rest.php","alexmar", "Jrtjeas33")
Applicants = applicantsModule()
a = session.get_entry(Applicants.module, )



#note = sugarcrm.Note(name="test note")
#session.set_entry(note)
#modules = session.get_available_modules()
#for m in modules:
#    print (m.module_key)
#note_query = sugarcrm.Note(name="test%")
#results = session.get_entry_list(note_query)
#print(results)
