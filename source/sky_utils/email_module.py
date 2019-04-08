import pdb
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def send_email(recipients, subject, message):

    fromEmail = 'sky@support.com'

    msg = MIMEMultipart()
    msg['From'] = fromEmail
    msg['To'] = ", ".join(recipients)
    msg['Subject'] = subject
    msg.attach(MIMEText(message))

    try:
        mailServer = smtplib.SMTP('192.168.100.220', 25)
        print('Connection to email server...')

        mailServer.send_message(msg)
        print('Email sent...')

        mailServer.quit()
        print('Connection closed...')

    except Exception as e:
        print(str(e))


def send_email_from(fEmail, recipients, subject, message, files):

    print(files)
    print(subject)
    print(message)

    files_temp = files
    fromEmail = fEmail+'@skyitgroup.com'
    attachment_name = files_temp.split('\\')[-1]
    msg = MIMEMultipart()
    msg['From'] = fromEmail
    msg['To'] = ", ".join(recipients)
    msg['Subject'] = subject
    msg.attach(MIMEText(message))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(files_temp, "rb").read())
    print("File opened to load attachment")
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename=' + attachment_name)
    msg.attach(part)
    print("attachment complete")

    try:
        mailServer = smtplib.SMTP('192.168.100.220', 25)
        print('Connection to email server...')

        mailServer.send_message(msg)
        print('Email sent...')

        mailServer.quit()
        print('Connection closed...')

    except Exception as e:
        print(str(e))


def send_email_test(recipients, subject, message):

    username = 'HPGroup@SkyITGroup.onmicrosoft.com'

    password =  'Vara2224'

    msg = MIMEMultipart()
    #msg['From'] = fromEmail
    msg['To'] = ", ".join(recipients)
    msg['Subject'] = subject
    msg.attach(MIMEText(message))

    try:
        mailServer = smtplib.SMTP('smtp-mail.outlook.com', 587)
        print('Connection to email server...')

        mailServer.ehlo()
        mailServer.starttls()
        mailServer.ehlo()
        mailServer.login(username,password)
        print('Logged in...')

        mailServer.sendmail(username, recipients, msg.as_string())
        print('Email sent...')

        mailServer.quit()
        print('Connection closed...')

    except Exception as e:
        print(str(e))



if __name__ == "__main__":
    import sys
    import os

    try:
        fromEmail = sys.argv[1]
        toEmail = sys.argv[2]
        subject = sys.argv[3]
        message = sys.argv[4]
        output = sys.argv[5]
        files =  output
        #main(fromEmail, [toEmail], subject, message,files)
    except IndexError:
        print('Please supply the following arguments: python email_notify.py <from_email_address> <to_email_address>')
        exit(1)  # Exit
    #send_email_from(fromEmail, [toEmail], 'CLS Style Selling Report', 'Please see attached.',files)
