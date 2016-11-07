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
    ws = wb[wbSheets[2]]

    vehicleTypeYear = (ws['A1'].value).split()
    build = OrderedDict([
        ("vehicleType", vehicleTypeYear[0]),
        ("year", vehicleTypeYear[1]),
        ("vehicleStyle", ws['I3'].value),
        ("description", ws['A3'].value),
        ("packages", [])
    ])

    flag = 1
    packageName = ""
    packagePrice = ""
    packageDetails = ""

    # Iteration of Available Packages
    for row in ws.rows:
        for cell in row:
            if 'CHOICE' in str(cell.value):
                packageName = cell.value
            if cell.column == 'F' and cell.value != None:
                if row[4].value == None:
                    packageDetails += row[6].value + "\n"
                else:
                    packageDetails += "%s ($%d)\n" % (row[6].value, row[9].value)
            if cell.column == 'I' and cell.value != None:
                packageNameCheck = packageName.lower()
                cellValueLow = cell.value.lower()
                if cellValueLow.startswith(packageNameCheck) and row[9].value:
                    packagePrice = row[9].value
                    package = {
                        "packageName": packageName,
                        "packagePrice": packagePrice,
                        "packageDetails": packageDetails
                    }
                    build["packages"].append(package)
                    packageDetails = ""

            # End of Available Packages, Prints Color Combo Starting Row
            if cell.value == 'AVAILABLE COLOR COMBINATIONS':
                print "Color Combinations starting at: (%s%s)" % (cell.column, cell.row)
                flag = 0
        if flag == 0:
            break

    print build

    with open("LexusGS.json", 'w') as outfile:
         json.dump(build, outfile, indent = 2)

if __name__ == '__main__':
    main()
