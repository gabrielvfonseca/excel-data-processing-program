# Path: /modules/progress.py
# Created: 2.11.2021
# Dev by Gabriel.

import sys
from time import sleep


class colors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    DEFAULT = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def progressBar(length, description):
    bar_length = 50

    for i in range(length+1):
        percent = 100.0*i/length
        sys.stdout.write('\r')
        sys.stdout.write("  {}: [{:{}}] {:>3}%".format(
            description, '='*int(percent/(100.0/bar_length)), bar_length, f'{colors.BOLD}{int(percent)}{colors.DEFAULT}'))
        sys.stdout.flush()
        sleep(0.02)
