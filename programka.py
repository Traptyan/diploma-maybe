import gspread
from oauth2client.client import SignedJwtAssertionCredentials
import pandas as pd
import json
import imaplib
import sugarcrm
from applicantsModule import applicantsModule
import time
import email
import re
import collections

while (True):
    user = 'bot.abiturientov@gmail.com'
    password = 'tcqnlzhcioqljofe'
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(user, password)

    mail.select('inbox')
    resp, unRead = mail.search(None, 'FROM', '"Form Notifications"', '(UNSEEN)')
    numMessages = len(unRead[0].split())
    print("new notifications: %s" % numMessages)

    items = unRead[0].split()

    filePaths = []
    for emailid in items:
        resp, data = mail.fetch(emailid, "(BODY[])")
        email_body = data[0][1].decode('utf-8')
        m = email.message_from_string(email_body)
        email_body = email_body.replace('=\r\n', '')
        filePaths.append(re.findall('\"https:\/\/docs\.google\.com\/spreadsheets\/d\/.*?\"', email_body)[0][1:-1])

    cnt = collections.Counter()
    for path in filePaths:
        cnt[path] += 1
    paths = cnt.keys()
    counts = cnt.values()

    for item in paths:
        SECRETS_FILE = "api-project-313216-1dd8217a1523.json"
        json_key = json.load(open(SECRETS_FILE))
        credentials = SignedJwtAssertionCredentials(json_key['client_email'],
                                                    json_key['private_key'], item)
        gc = gspread.service_account()
        workbook = gc.open_by_url(item)
        sheet = workbook.sheet1
        data = pd.DataFrame(sheet.get_all_values())

        column_names = {'Отметка времени': 'Timestamp',
                        'ФИО': 'first_name',
                        'E-mail': 'E-mail',
                        'Мобильный телефон': 'phone_mobile',
                        'Город': 'city',
                        'Школа': 'school',
                        'Класс (ТОЛЬКО ЦИФРА)': 'grade',
                        }
        data.rename(columns=column_names, inplace=True)
        convertedArray = data.to_numpy()
        print(convertedArray)

        session = sugarcrm.Session("https://crm.uni-dubna.ru/service/v4/rest.php", "Botich", "D1pomchik_Rul1t")
        for j in range(cnt[item]):
            length = len(convertedArray)
            Applicants = applicantsModule(first_name=convertedArray[length-(j+1), 1],
                                          email1=convertedArray[length-(j+1), 2],
                                          phone_mobile=convertedArray[length-(j+1), 3],
                                          primary_address_city=convertedArray[length-(j+1), 4],
                                          school=convertedArray[length-(j+1), 5],
                                          grade_c=convertedArray[length - (j + 1), 6],
                                          event=sheet.spreadsheet.title.split(' ')[0],
                                          subject=sheet.spreadsheet.title.split(' ')[1])
            #session.set_entry(Applicants)

    time.sleep(60)