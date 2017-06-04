import win32com.client
import pandas as pd

file = pd.ExcelFile('Insert directory of Excel file containing list of email addresses you plan to send')

df = file.parse('Sheet1')

for index, row in df.iterrows():
    email = (row['Email Address'])
    subject = (row['Subject'])
    body = str((row['Email HTML Body']))

    if (pd.isnull(email) or pd.isnull(subject) or pd.isnull(body)):
        continue

    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = subject
    newMail.HTMLbody = (r"" +
    body +
    "")

    newMail.To = email
    newMail.display()