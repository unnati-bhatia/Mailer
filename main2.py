import openpyxl as xl
import base64
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (Mail, Attachment, FileContent, FileName, FileType, Disposition)

wb = xl.load_workbook(r'C:\Users\Dell\Documents\emailer.xlsx')
sheet1 = wb["Sheet1"]

names = []
contacts = []
images = []

for cell in sheet1['A']:
    contacts.append(cell.value)

for cell in sheet1['B']:
    names.append(cell.value)

for cell in sheet1['C']:
    images.append(cell.value)

for i in range(len(contacts)):
    message = Mail(from_email='(abc)', #to be written
                   to_emails=contacts[i],
                   subject='Subject',
                   html_content= '''
    (html file here) ''')

    with open(images[i], 'rb') as f:
        data = f.read()
        f.close()
    encoded_file = base64.b64encode(data).decode()

    attachedFile = Attachment(
        FileContent(encoded_file),
        FileName(names[i]),
        FileType('application/pdf'),
        Disposition('attachment')
    )
    message.attachment = attachedFile

    sg = SendGridAPIClient('API Key') #to be generated
    response = sg.send(message)
    print(images[i] + ' done')

print("\nAll done")
