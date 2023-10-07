# Path: /modules/excel.py
# Created: 31.11.2021
# Dev by Gabriel.

import sys
import os
import time
import openpyxl
from urllib import request
from time import sleep

# Import Modules
from ..main import askInput, colors
from .progress import progressBar as pg
from .table import designTable
from .graph import drawGraphic


def getNewFile():
    def getFileFromWeb(url, local):
        request.urlretrieve(url, local)
        file = os.stat(str(local))
        pg.progressBar(file.st_size, str(local)[:-5])

    print(f"\nDownload file from:")

    local = "", "new_spreadsheet.xlsx"  # Set by default

    while True:
        new_input = askInput()

        if (new_input == "exit"):
            # Terminate Program
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
            time.sleep(2)
            break


def readFile(self):
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
