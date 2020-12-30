# Code to create a xslx file from json (ID).
# Install openpyxl pip3 install openpyxl


import json
import openpyxl
from openpyxl import Workbook


folder = 'Caliente26Nov-27Dic'
fileJson = f'./json/{folder}/id.json'
fileXslx = f'./excel/{folder}/id.xlsx'


def set_id():
    if 'id' in idR:
        ws_01.cell(row, 1, idR["id"])
    else:
        ws_01.cell(row, 1, "")


def set_checkTime():
    if 'checkTime' in idR:
        ws_01.cell(row, 4, idR["checkTime"])
    else:
        ws_01.cell(row, 4, "")


def set_documentType():
    if 'documentType' in idR:
        ws_01.cell(row, 5, idR["documentType"])
    else:
        ws_01.cell(row, 5, "")


def set_countryCode():
    if 'countryCode' in idR:
        ws_01.cell(row, 6, idR["countryCode"])
    else:
        ws_01.cell(row, 6, "")


def set_spentCredits():
    if 'spentCredits' in idR:
        ws_01.cell(row, 7, idR["spentCredits"])
    else:
        ws_01.cell(row, 7, "")


def set_totalChecks():
    if 'totalChecks' in idR:
        ws_01.cell(row, 9, idR["totalChecks"])
    else:
        ws_01.cell(row, 9, "")


def set_successChecks():
    if 'successChecks' in idR:
        ws_01.cell(row, 10, idR["successChecks"])
    else:
        ws_01.cell(row, 10, "")


def set_warningChecks():
    if 'warningChecks' in idR:
        ws_01.cell(row, 11, idR["warningChecks"])
    else:
        ws_01.cell(row, 11, "")


def set_failedChecks():
    if 'failedChecks' in idR:
        ws_01.cell(row, 12, idR["failedChecks"])
    else:
        ws_01.cell(row, 12, "")


if __name__ == '__main__':

    json_data = {}

    with open(fileJson) as json_file:
        json_data = json.load(json_file)

    wb = Workbook()

    # Grab the active worksheet
    ws_01 = wb.active

    # Set the title of the worksheet
    ws_01.title = 'ID Records'

    # Set first row
    ws_01.cell(1, 1, "ID")
    ws_01.cell(1, 2, "Identifier")
    ws_01.cell(1, 3, "Start Date Utc")
    ws_01.cell(1, 4, "Check Time")
    ws_01.cell(1, 5, "Document Type")
    ws_01.cell(1, 6, "Country Code")
    ws_01.cell(1, 7, "Spent Credits")
    ws_01.cell(1, 8, "Status")
    ws_01.cell(1, 9, "Total Checks")
    ws_01.cell(1, 10, "Success Checks")
    ws_01.cell(1, 11, "Warning Checks")
    ws_01.cell(1, 12, "Failed Checks")

    row = 1
    for idR in json_data.get("verifications"):
        row += 1
        set_id()
        ws_01.cell(row, 2, idR["identifier"])
        ws_01.cell(row, 3, idR["startDateUtc"])
        set_checkTime()
        set_documentType()
        set_countryCode()
        set_spentCredits()
        ws_01.cell(row, 8, idR["status"])
        set_totalChecks()
        set_successChecks()
        set_warningChecks()
        set_failedChecks()

    # Save it in an Excel file
    wb.save(fileXslx)
