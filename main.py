# Path: /main.py
# Created: 31.11.2021
# Dev by Gabriel.

import os
import sys
import pkg_resources
from time import sleep

# Import Modules
# from modules.excel import getNewFile, readFile
# from modules.smtp import sendEmail


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


username = str(input("\nEnter your name to start: "))
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
