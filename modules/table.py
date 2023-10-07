# Path: /modules/table.py
# Created: 3.11.2021
# Dev by Gabriel.

from prettytable import PrettyTable

def designTable(heading, rows):
    table = PrettyTable(heading)  # Table Collums
    i = 0

    # print(f"Number of Rows: {len(rows)}")

    for i in range(len(rows)):
        table.add_row(rows[i])

    return table
