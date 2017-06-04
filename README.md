# python-outlook-mail-merge
This project allows you send mass emails with rich HTML formatting through Microsoft Outlook via an Excel workbook and Python.

# Getting Started
You will need to install the libaries pypiwin32 (win32com.client) and pandas libraries to get this project running on your local machine.
```
pip install pypiwin32
```
```
pip install pandas
```

# Running the Script
Create an Excel workbook with the column headers 'Email Address', 'Subject', and 'HTML Email Body'. Insert the email addresses, subject lines,
and HTML email bodies in the respective cells under the column headers.

Replace the code in the following line of the python script with the file directory to your Excel workbook:
```
file = pd.ExcelFile('Insert directory of Excel file containing list of email addresses you plan to send') 
```
# License
See the LICENSE file for license rights and limitations (MIT).
