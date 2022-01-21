import openpyxl as xl
import smtplib
#import imghdr
from email.message import EmailMessage

EMAIL_ADDRESS = '(abc)' #to be written
EMAIL_PASSWORD = '(abc)'  #to be written

wb = xl.load_workbook(r'C:\Users\Dell\Documents\emailer.xlsx')
sheet1 = wb["Sheet1"]

names = []
contacts = []
files = []

for cell in sheet1['A']:
    contacts.append(cell.value)

for cell in sheet1['B']:
    names.append(cell.value)

for cell in sheet1['C']:
    files.append(cell.value)


for i in range(len(contacts)):
    msg = EmailMessage()
    msg['Subject'] = 'Subject Title'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = contacts[i]
    with open(files[i], 'rb') as f:
        file_data = f.read()
        #file_type = imghdr.what(f.name)
        file_name = names[i]
        
    msg.set_content('''
    xyz''',  subtype='html')
    
    msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)
    print("Mail sent to " + files[i])

print("All mails sent!")


