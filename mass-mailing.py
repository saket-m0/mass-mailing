# Importing the Libraries
import smtplib
from email.message import EmailMessage
import imghdr
import pandas as pd
from credentials import EMAIL_ADDRESS, PASSWORD
from mailer import MAILER, ALTERNATIVE

def MakeMessage(SendTo):
    # Setting up the Message
    MESSAGE = EmailMessage()
    MESSAGE['Subject'] = '[Business Standard] : Digital Newspaper Subscription for Students'
    MESSAGE['From'] = EMAIL_ADDRESS
    MESSAGE['Bcc'] = SendTo
    MESSAGE.set_content('[Business Standard] : Digital Newspaper Subscription for Students')
    MESSAGE.add_alternative(MAILER, subtype='html')

    # Attaching Images to mail
    images = ['img.jpg']

    for image in images:
        with open(image, 'rb') as file:
            file_data = file.read()
            file_type = imghdr.what(file.name)
            file_name = file.name

        MESSAGE.add_attachment(file_data, maintype='image', subtype=file_type, filename=file_name)
        
    # Attaching Pdfs to mail
    pdfs = []

    for pdf in pdfs:
        with open(pdf, 'rb') as file:
            file_data = file.read()
            file_name = file.name

        MESSAGE.add_attachment(file_data, maintype='application', subtype='octate-stream', filename=file_name) 

    return MESSAGE


def SendMail(MESSAGE, Email, Password):
    # Function to send mails
    # Connecting to Gmail using SSL
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(Email, Password) #Logging in to gmail
        smtp.send_message(MESSAGE) #Sending the mail

# Setting Mailing List
Data = pd.read_excel(r'MailingList.xlsx') # Reading Excel File
MailingList = Data['Email Address'].tolist()  # Extracting Email Addresses and putting in a list

MailIds = [
    ['manas.dedhia@ecell-iitkgp.org', 'ecell1234'],
    ['adnan.hassan@ecell-iitkgp.org', 'Ecell1234'],
    ['prathamesh.ugale@ecell-iitkgp.org', 'Ecell1234']
]

count = 0

for i in range(len(MailIds)):
    for j in range(10):
        MailTo = []
        for k in range(40):
            if(count < len(MailingList)):
                MailTo.append(MailingList[count])
                count += 1
        if (len(MailTo) != 0):
            Message = MakeMessage(MailTo)
            SendMail(Message, MailIds[i][0], MailIds[i][1])