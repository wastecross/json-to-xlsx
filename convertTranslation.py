# Code to create a xslx file from json (INE).
# Install openpyxl pip3 install openpyxl


import json
import openpyxl
from openpyxl import Workbook


folder = 'translations'
file = 'api'
fileJson = f'./json/{folder}/{file}.json'
fileXslx = f'./excel/{folder}/{file}.xlsx'


def set_language():
    if 'en' in translationRecord:
        ws_01.cell(row, 1, translationRecord["en"])
    else:
        ws_01.cell(row, 1, "")


def set_data():
    if 'data' in translationRecord:
        for attr, value in translationRecord["data"].items():
            if attr == 'successChecks':
                ws_01.cell(row, 3, value)
            if attr == 'warningChecks':
                ws_01.cell(row, 4, value)
            if attr == 'failedChecks':
                ws_01.cell(row, 5, value)
            if attr == 'globalResult':
                ws_01.cell(row, 6, value)
    else:
        ws_01.cell(row, 3, "")
        ws_01.cell(row, 4, "")
        ws_01.cell(row, 5, "")
        ws_01.cell(row, 6, "")


if __name__ == '__main__':

    json_data = {}

    with open(fileJson) as json_file:
        json_data = json.load(json_file)

    wb = Workbook()

    # Grab the active worksheet
    ws_01 = wb.active

    # Set the title of the worksheet
    ws_01.title = 'Translation'

    # Set first row
    ws_01.cell(1, 1, "English")
    ws_01.cell(1, 2, "Spanish")

    row = 1
    for translationRecord in json_data.get("languages"):
        row += 1
        set_language()

    # Save it in an Excel file
    wb.save(fileXslx)
