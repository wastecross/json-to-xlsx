# Code to create a xslx file from json (INE).
# Install openpyxl pip3 install openpyxl


import json
import openpyxl
from openpyxl import Workbook


folder = 'Nexu01-31Dic'
fileJson = f'./json/{folder}/ine.json'
fileXslx = f'./excel/{folder}/ine.xlsx'


def set_id():
    if 'id' in ineR:
        ws_01.cell(row, 1, ineR["id"])
    else:
        ws_01.cell(row, 1, "")


def set_createdAt():
    if 'createdAt' in ineR:
        ws_01.cell(row, 2, ineR["createdAt"])
    else:
        ws_01.cell(row, 2, "")


if __name__ == '__main__':

    json_data = {}

    with open(fileJson) as json_file:
        json_data = json.load(json_file)

    wb = Workbook()

    # Grab the active worksheet
    ws_01 = wb.active

    # Set the title of the worksheet
    ws_01.title = 'INE Records'

    # Set first row
    ws_01.cell(1, 1, "ID")
    ws_01.cell(1, 2, "Created At")

    row = 1
    for ineR in json_data.get("verifications"):
        row += 1
        set_id()
        set_createdAt()

    # Save it in an Excel file
    wb.save(fileXslx)
