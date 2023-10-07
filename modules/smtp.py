# Path: /modules/smtp.py
# Created: 6.11.2021
# Dev by Gabriel.

import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

# Import Modules
from ..main import username  # Get username

load_dotenv()


def sendEmail():
    port = int(os.getenv('SSL_PORT'))  # For SSL
    sender_address = str(os.getenv('SSL_PORT'))  # Sender Email
    sender_password = input("Enter email password: ")  # Password
    email_subject = "Eleições Autárquicas"  # Subject
    receiver_address = str(os.getenv('RECEIVER_EMAIL'))  # Receiver Email
    smtp_server = str(os.getenv('SMTP_SERVER'))  # SMTP Email Server
    attach_file_name = 'spreadsheetpdf'  # File attached

    body = f"""\
    Exmºos Srs da CNE

    Seguem em anexo os arquivos com os respetivos dados eleitorais, bem como as respetivas mudificações conforme solicitados.

    Atenciosamente 
    {username}
    """

    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = email_subject

    # The subject line
    # The body and the attachments for the mail
    message.attach(MIMEText(body, 'plain'))
    attach_file = open(attach_file_name, 'rb')  # Open the file as binary mode
    payload = MIMEBase('application', 'octate-stream')
    payload.set_payload((attach_file).read())
    encoders.encode_base64(payload)  # Encode the attachment
    payload.add_header('Content-Decomposition', 'attachment',
                       filename=attach_file_name)
    message.attach(payload)

    # Create SMTP session for sending the mail
    session = smtplib.SMTP(smtp_server, port)
    session.starttls()  # Enable security

    # Login with mail_id and password
    session.login(sender_address, sender_password)
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)
    session.quit()

    print('Mail sent successfully!')
