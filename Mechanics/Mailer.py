'''
Created on Jan 25, 2018

@author: James

===========================================================
Thanks to http://naelshiab.com/tutorial-send-email-python/
for the short and easy explanation of making an E-mail bot!
===========================================================
'''

import smtplib

# Multipurpose Internet Mail Extensions (MIME)
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def Mailer():
    file = open("text.txt" , 'r')
    whatIsThis = file.readline()
    sendingAddress = file.readline()
    recievingAddress = file.readline()
    file.close
    
    # Breaks Email into multiple parts allowing detailed info such as a subject (Like a descriptor, or property)
    message = MIMEMultipart()
    
    message['From'] = sendingAddress
    message['To'] = recievingAddress
    message['Subject'] = 'Daily Numbers'
    
    # Messages body text
    body = 'See attachments'
    
    # Attaches body to message in plain format
    message.attach(MIMEText(body, 'plain'))
    
    # Setups Excel file for attachment
    workbook = 'Log.xlsx'
    attachment = open(workbook, 'rb')
    part = MIMEBase('application', 'octet-stream')
    
    '''
    ======================================================================================= 
    octet-stream is a binary file
    
    This is the default value for a binary file. As it really means unknown binary file,
    browsers usually don't automatically execute it, or even ask if it should be executed.
    They treat it as if the Content-Disposition header was set with the value attachment
    and propose a 'Save As' file.
    ======================================================================================= 
    ---------------------------------------------------------------------------------------
    THANKS TOO: https://developer.mozilla.org/en-US/docs/Web/HTTP/Basics_of_HTTP/MIME_types
    ---------------------------------------------------------------------------------------
    '''
    
    # Encodes and adds attachment information
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= Log.xlsx")
    
    # Attaches attachment 
    message.attach(part)
    
    # Makes secure connection to gmail servers 
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sendingAddress, whatIsThis)
    # Converts message "Object" to string
    text = message.as_string()
    # Sends mail with actual email address as sender (spoofing is bad!) and disconnect from gmails server.
    server.sendmail(sendingAddress, recievingAddress, text)
    server.quit()



