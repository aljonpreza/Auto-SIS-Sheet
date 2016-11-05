from openpyxl import load_workbook
from collections import OrderedDict
import os
import sys
import requests
import json
import argparse

def main():
    wb = load_workbook(filename='SISSheet.xlsx', data_only = True)

    # Array of all sheets in spreadsheet
    wbSheets = wb.get_sheet_names()
    ws = wb[wbSheets[1]]

    vehicleTypeYear = (ws['A1'].value).split()
    build = OrderedDict([
        ("vehicleType", vehicleTypeYear[0]),
        ("year", vehicleTypeYear[1]),
        ("vehicleStyle", ws['I3'].value),
        ("description", ws['A3'].value)
    ])

    flag = 1
    for row in ws.rows:
        for cell in row:
            if 'CHOICE' in str(cell.value):
                print "%s: <%s%s>" % (cell.value, cell.column, cell.row)
            if cell.column == 'F' and cell.value != None:
                print cell.value
            if cell.value == 'AVAILABLE COLOR COMBINATIONS':
                print "Color Combinations starting at: (%s%s)" % (cell.column, cell.row)
                flag = 0
        if flag == 0:
            break

    # with open("LexusGS.json", 'w') as outfile:
    #     json.dump(build, outfile, indent = 2)

if __name__ == '__main__':
    main()
