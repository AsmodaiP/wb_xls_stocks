import logging
import openpyxl
from googleapiclient.discovery import build
from google.oauth2 import service_account
import os
from dotenv import load_dotenv
from pycel.excelcompiler import ExcelCompiler

load_dotenv()
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = os.environ['SPREADSHEET_ID']

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'credentials_service.json')
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)


def get_data(filename='stocks.xlsx'):
    wb = openpyxl.load_workbook(filename, data_only=False)
    excel = ExcelCompiler(filename='stocks.xlsx')
    list_number = 0
    data = {}
    for sheetname in wb.sheetnames:
        if sheetname == 'Зеркала':
            continue
        print(f'{sheetname}!5')
        wb.active = list_number
        list_number += 1
        ws = wb.active

        row_number = 1
        for row in ws.rows:
            barcode, article, count = (
                row[0].value, row[1].value, excel.evaluate(f'{sheetname}!F{row_number}'))
            if row[0].value is not None:
                data[barcode] = {'article': article, 'count': count}
            row_number += 1
    return data


def get_body_data(data, range_name='Баркоды'):
    body_data = []
    i = 1
    for barcode, info in data.items():
        article = info['article']
        count = info['count']
        body_data += [{'range': f'{range_name}!A{i}', 'values': [[barcode]]},
                      {'range': f'{range_name}!B{i}', 'values': [[article]]},
                      {'range': f'{range_name}!C{i}', 'values': [[count]]},
                      ]
        i += 1
    logging.info((barcode, article, count))
    return body_data


def update_google_sheet(spreadsheet_id=SPREADSHEET_ID, range_name='Баркоды'):
    data = get_data()
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    body_data = get_body_data(data, range_name)
    body = {
        'valueInputOption': 'USER_ENTERED',
        'data': body_data}

    sheet.values().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
