# Path: /all-in-one.py
# Dev by Gabriel.

import os
import sys
import ssl
import email
import smtplib
import openpyxl
import numpy as np
import pkg_resources
from time import sleep
from urllib import request
from email import encoders
from dotenv import load_dotenv
import matplotlib.pyplot as plt
from prettytable import PrettyTable
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class colors:
    HEADER = "\033[95m"
    OKBLUE = "\033[94m"
    OKCYAN = "\033[96m"
    OKGREEN = "\033[92m"
    WARNING = "\033[93m"
    FAIL = "\033[91m"
    DEFAULT = "\033[0m"
    BOLD = "\033[1m"
    UNDERLINE = "\033[4m"


def askInput():
    return input((colors.OKGREEN + "\n{}@python: ".format(username)) + colors.DEFAULT)


def designTable(heading, rows):
    table = PrettyTable(heading)  # Table Collums
    i = 0

    # print(f"Number of Rows: {len(rows)}")

    for i in range(len(rows)):
        table.add_row(rows[i])

    return table


def sendEmail():
    port = int(os.getenv('SSL_PORT'))  # For SSL
    sender_address = str(os.getenv('SSL_PORT'))  # Sender Email
    sender_password = input("Enter email password: ")  # Password
    email_subject = "Eleições Autárquicas"  # Subject
    receiver_address = str(os.getenv('RECEIVER_EMAIL'))  # Receiver Email
    smtp_server = str(os.getenv('SMTP_SERVER'))  # SMTP Email Server
    attach_file_name = 'spreadsheet.xlsx'  # File attached

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


def progressBar(length, description):
    bar_length = 50

    for i in range(length+1):
        percent = 100.0*i/length
        sys.stdout.write('\r')
        sys.stdout.write("  {}: [{:{}}] {:>3}%".format(
            description, '='*int(percent/(100.0/bar_length)), bar_length, f'{colors.BOLD}{int(percent)}{colors.DEFAULT}'))
        sys.stdout.flush()


def drawGraphic(x, y):
    plt.plot(x, y)
    plt.show()


def getNewFile():
    def getFileFromWeb(url, local):
        request.urlretrieve(url, local)
        file = os.stat(str(local))
        progressBar(file.st_size, str(local)[:-5])

    local = "spreadsheet.xlsx"  # Set by default

    while True:
        new_input = askInput()

        if (new_input == "exit"):
            # Terminate Program
            print((colors.HEADER + "Shutting down...\n") + colors.DEFAULT)
            sleep(2)
            sys.exit()

        elif (new_input == "help"):
            # Display avaiable commands
            print("\nList of available program commands:")
            head = ["Command", "Description"]
            row = [["'Enter'", "- Input new online file destiny"],
                   ["exit", "- Stop program"],
                   ["help", "- Show available commands and descriptions"]]

            print(designTable(head, row), "\n")

        else:
            getFileFromWeb(input("  Download file from: "), local)
            sleep(2)
            break


def readFile():
    file = "spreadsheet.xlsx"
    new_file = "dados_eleicao_final.xlsx"
    wookbook = openpyxl.load_workbook(file)
    sheet = wookbook.active

    number = 1
    letter = "B"
    excel = [
        [], [], [], []
    ]

    for i in range(len(excel)):
        number += 1
        for x in range(5):
            coordenate = (letter + str(number))
            value = int(sheet[coordenate].value)
            excel[i].append(value)

            letter = chr((ord(letter)) + 1)
            if (letter == "G"):
                letter = "B"

    # Calculate Vertical Values
    soma = 0
    vertical = []

    for i in range(len(excel)):
        for numb in excel[i]:
            soma += numb
        vertical.append(soma)
        soma = 0

    # Calculate Collum Values
    soma = 0
    collum = []

    for i in range(len(excel)):
        for x in range(4):
            soma += excel[x][i]
        collum.append(soma)
        soma = 0

    print(collum)

    # Update Total
    sheet['G1'] = "Totais"
    sheet['G2'] = vertical[0]
    sheet['G3'] = vertical[1]
    sheet['G4'] = vertical[2]
    sheet['G5'] = vertical[3]

    # Update %

    def percentage(part, whole):
        return str(100 * float(part)/float(whole)) + "%"

    sheet['H1'] = "% Votos"
    sheet['H2'] = percentage(vertical[0], 5)
    sheet['H3'] = percentage(vertical[1], 5)
    sheet['H4'] = percentage(vertical[2], 5)
    sheet['H5'] = percentage(vertical[3], 5)

    # Max & Min
    sheet['A6'] = "Max"
    sheet['A7'] = "Min"
    maximum = []
    minium = []

    m = 0
    letter = "B"
    for m in range(len(collum)):
        temp = int(max(collum[m]))
        sheet[str(chr((ord(letter)) + 1) + "6")] = temp
        maximum.append[temp]

    n = 0
    for n in range(len(collum)):
        temp = int(max(collum[n]))
        sheet[str(chr((ord(letter)) + 1) + "7")] = int(min(collum[n]))
        minium.append[temp]

    head = ["Freguesia/Partido", 1, 2, 3, 4, 5, "Totais", "% Votos"]
    body = [
        ["A", vertical[0]], ["B", vertical[1]], ["C", vertical[2]], [
            "D", vertical[3]], ["Max", maximum], ["Min", minium]
    ]

    print(designTable(head, body))  # Design Table
    drawGraphic(["A", "B", "C", "D"], [vertical[0], vertical[1],
                vertical[2], vertical[3]])  # Draw Graphic

    sheet.save(new_file)  # Save as new file with a different name


username = (str(input("\nEnter your name to start: "))).replace(" ", "-")
sleep(1)

print(
    r"""
 ____              __                            ___      
/\  _`\           /\ \               __         /\_ \     
\ \ \L\_\     __  \ \ \____   _ __  /\_\      __\//\ \    
 \ \ \L_L   /'__`\ \ \ '__`\ /\`'__\\/\ \   /'__`\\ \ \   
  \ \ \/, \/\ \L\.\_\ \ \L\ \\ \ \/  \ \ \ /\  __/ \_\ \_ 
   \ \____/\ \__/.\_\\ \_,__/ \ \_\   \ \_\\ \____\/\____\
    \/___/  \/__/\/_/ \/___/   \/_/    \/_/ \/____/\/____/
"""
)

print("Dev by Gabriel.")
print(
    """
This python program allows through a homemade computer terminal
to get from web a specific excel file (.xslx), read collums and 
lines and do calculations using does values get from the file.
At the end, data must be written on a new file and sent by email 
to reciver.
    """
)
print("Packages instalation status: \n")

i, dot = 0, ""

# Get librarys from txt
packages = open("requirements.txt")
line = packages.read().replace("\n", " ")
packages.close()
line = line.split(" ")
required = set(line)
installed = {pkg.key for pkg in pkg_resources.working_set}

missing = required - installed
ready = required - missing

for item in ready:
    for i in range(34 - len(item)):
        dot = "{}{}".format(dot, ".")
        i = +1
    print(f"{item} {dot} [{colors.OKGREEN}True{colors.DEFAULT}]")
    dot = ""

for item in missing:
    for i in range(34 - len(item)):
        dot = "{}{}".format(dot, ".")
        i = +1
    print(f"{item} {dot} [{colors.FAIL}False{colors.DEFAULT}]")
    dot = ""


getNewFile()  # Request user a file to download
readFile()  # Read file and calculate data
sendEmail()  # Send email
