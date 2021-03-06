import gspread
from oauth2client.client import SignedJwtAssertionCredentials
import pandas as pd
import json
import imaplib
from imap_tools import MailBox, AND
import sugarcrm
from applicantsModule import applicantsModule

SCOPE = "https://docs.google.com/spreadsheets/d/1RmKzWzTIaVn2XsQKcupycPQWfHvOrWIkENv9VGK-1L4/edit#gid=1467420316"

#SCOPE = ["https://drive.google.com/drive/my-drive"]
SECRETS_FILE = "api-project-313216-1dd8217a1523.json"
SPREADSHEET = "Новая форма (Ответы)"
json_key = json.load(open(SECRETS_FILE))
credentials = SignedJwtAssertionCredentials(json_key['client_email'],
                                            json_key['private_key'], SCOPE)
gc = gspread.service_account()
for sheet in gc.openall():
    print("{} - {}".format(sheet.title, sheet.id))
workbook = gc.open_by_url(SCOPE)
sheet = workbook.sheet1
data = pd.DataFrame(sheet.get_all_values())
print(data)


user = 'alexandr.martinovitch@gmail.com'
password = 'fgleztpdvrhiifvv'
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(user, password)


mail.select('inbox')
unRead = mail.uid('Search', 'FROM', '"me"', '(UNSEEN)')
numMessages = len(unRead[1][0].split())
print("messages from me: %s" % numMessages)



with MailBox('imap.gmail.com').login(user,password) as mailbox:
    anred = mailbox.fetch("FROM me UNSEEN")
    mailbox.seen(anred, False) #false развидел

unRead = mail.uid('Search', 'FROM', '"me"', '(UNSEEN)')
numMessages = len(unRead[1][0].split())
print("messages from me: %s" % numMessages)

'''
column_names = {'Отметка времени': 'Timestamp',
                'ФИО': 'first_name',
                'Телефон': 'phone_mobile',
}
data.rename(columns=column_names, inplace=True)

convertedArray = data.to_numpy()

session = sugarcrm.Session("https://crm.uni-dubna.ru/service/v4/rest.php", "alexmar", "Jrtjeas33")
print(len(convertedArray))
for i in range(len(convertedArray)-1):
    Applicants = applicantsModule(first_name=convertedArray[i,1], phone_mobile=convertedArray[i,2])
    session.set_entry(Applicants)

'''




'''note = sugarcrm.Note(name="test note")
session.set_entry(note)
modules = session.get_available_modules()
for m in modules:
    print(m.module_key)
note_query = sugarcrm.Note(name="test%")
results = session.get_entry_list(note_query)
print(results)'''
