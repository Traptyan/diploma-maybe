import gspread
from oauth2client.client import SignedJwtAssertionCredentials
import pandas as pd
import json
import imaplib
from imap_tools import MailBox, AND
import sugarcrm
from applicantsModule import applicantsModule
import time
import email
import re


while(True):

    user = 'bot.abiturientov@gmail.com'
    password = 'tcqnlzhcioqljofe' # D1pomchik_Rul1t
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(user, password)

    mail.select('inbox')
    resp, unRead = mail.search(None, 'FROM', '"Form Notifications"', '(UNSEEN)')
    numMessages = len(unRead[0].split())
    print("messages from me: %s" % numMessages)

    items = unRead[0].split()
    for emailid in items:
        resp, data=mail.fetch(emailid, "(BODY[])")
        email_body=data[0][1].decode('utf-8')
        m=email.message_from_string(email_body)

        filePath=''
        for part in m.walk():
            email_body=email_body.replace('=\r\n','')
            filePath=re.findall('\"https:\/\/docs\.google\.com\/spreadsheets\/d\/.*?\"',email_body)[0][1:-1]

        SECRETS_FILE = "api-project-313216-1dd8217a1523.json"
        json_key = json.load(open(SECRETS_FILE))
        credentials = SignedJwtAssertionCredentials(json_key['client_email'],
                                                    json_key['private_key'], filePath)
        gc = gspread.service_account()
        for sheet in gc.openall():
            print("{} - {}".format(sheet.title, sheet.id))
        workbook = gc.open_by_url(filePath)
        sheet = workbook.sheet1
        data = pd.DataFrame(sheet.get_all_values())
        print(data)


        column_names = {'Отметка времени': 'Timestamp',
                        'ФИО': 'first_name',
                        'Телефон': 'phone_mobile',
                        }
        data.rename(columns=column_names, inplace=True)

        convertedArray = data.to_numpy()

        session = sugarcrm.Session("https://crm.uni-dubna.ru/service/v4/rest.php", "alexmar", "Jrtjeas33")
        print(len(convertedArray))
        for i in range(len(convertedArray) - 1):
            Applicants = applicantsModule(first_name=convertedArray[i, 1], phone_mobile=convertedArray[i, 2])
            session.set_entry(Applicants)

    time.sleep(60)