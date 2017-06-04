import win32com.client
import pandas as pd

file = pd.ExcelFile('Insert directory of Excel file containing list of email addresses you plan to send') #Establishes the excel file you wish to import into Pandas

df = file.parse('Sheet1') #Uploads Sheet1 from the Excel file into a dataframe

for index, row in df.iterrows(): #Loops through each row in the dataframe
    email = (row['Email Address'])  #Sets dataframe variable, 'email' to cells in column 'Email Addresss'
    subject = (row['Subject']) #Sets dataframe variable, 'subject' to cells in column 'Subject'
    body = str((row['Email HTML Body'])) #Sets dataframe variable, 'body' to cells in column 'Email HTML Body'

    if (pd.isnull(email) or pd.isnull(subject) or pd.isnull(body)): #Skips over rows where one of the cells in the three main columns is blank
        continue

    olMailItem = 0x0 #Initiates the mail item object
    obj = win32com.client.Dispatch("Outlook.Application") #Initiates the Outlook application
    newMail = obj.CreateItem(olMailItem) #Creates an Outlook mail item
    newMail.Subject = subject #Sets the mail's subject to the 'subject' variable
    newMail.HTMLbody = (r"" +
    body +
    "") #Sets the mail's body to 'body' variable

    newMail.To = email #Sets the mail's To email address to the 'email' variable
    newMail.display() #Displays the mail as a draft email