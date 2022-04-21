
import imp
import logging
from typing import List

import requests
from datetime import datetime
import json
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv
import telegram

from sheet import get_counts_from_table




SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'credentials_service.json')
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
START_POSITION_FOR_PLACE = 14

FIRST_INDEX = 6

load_dotenv()
SPREADSHEET_ID_STOCKS = os.environ.get('SPREADSHEET_ID_FOR_STOCKS')
SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID')

def get_sheetnames(spreadsheet_id):
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    sheets = sheet.get(spreadsheetId=spreadsheet_id).execute().get('sheets')
    sheetnames = []
    for sheet in sheets:
        sheetnames.append(sheet.get("properties", {}).get("title", "Sheet1"))
    return sheetnames

def get_sum_by_barcode(mirrors_barcodes, data, row, first_index=FIRST_INDEX):
    sum_of_mirrors = 0
    barcode = row[0]
    if barcode in mirrors_barcodes:
        for bar in mirrors_barcodes:
            if str(bar) in data:
                sum_of_mirrors += int(data[str(bar)])
                del data[str(bar)]
    else:
        sum_of_mirrors = int(data[barcode])
        del data[str(barcode)]
    print(row)
    logging.info(row)
    today_sells = row[first_index + datetime.now().day-1]
    if today_sells is not None and today_sells  != '':
        return sum([int(today_sells), sum_of_mirrors])
    return sum_of_mirrors

def convert_to_column_letter(column_number):
    column_letter = ''
    while column_number != 0:
        c = ((column_number - 1) % 26)
        column_letter = chr(c + 65) + column_letter
        column_number = (column_number - c) // 26
    return column_letter

def get_mirrors(spreadsheet_id) -> List[List[str]]:
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                range=f'Зеркала!A:A', majorDimension='ROWS').execute()
    return (result.get('values', []))

def insert_data_in_table(data, spreadsheet_id, first_index=FIRST_INDEX):
    '''inser data in google table'''
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    sheetnames = get_sheetnames(spreadsheet_id)
    mirrors = get_mirrors(spreadsheet_id)
    body_data = []
    for sheetname in sheetnames:
        if sheetname == 'Зеркала':
            continue
        result = sheet.values().get(spreadsheetId=spreadsheet_id,
                            range=f'{sheetname}!A:BS', majorDimension='ROWS').execute()
        values = result.get('values', [])
        i = 1
        for row in values:
            logging.info(row)
            try:
                bar = row[0]
            except:
                i+=1
                continue
            if bar in data:
                new_value = get_sum_by_barcode(mirrors, data, row, first_index) 
                print(new_value)
                body_data.append([{'range': f'{sheetname}!{convert_to_column_letter(first_index + datetime.now().day)}{i}', 'values': [[new_value]]}])
            i+=1
        body = {
        'valueInputOption': 'USER_ENTERED',
        'data': body_data}
    if data:
        return {'erorrs': data}
    sheet.values().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return {'erorrs': data}

        
def insert_sales(filename):
    return insert_data_in_table(
        get_counts_from_table(filename), first_index=FIRST_INDEX, spreadsheet_id=SPREADSHEET_ID_STOCKS)

def insert_supplie(filename):
    return insert_data_in_table(
        get_counts_from_table(filename), first_index=FIRST_INDEX+33, spreadsheet_id=SPREADSHEET_ID_STOCKS) 


def get_all_data_from_google_sheet(spreadsheet_id=SPREADSHEET_ID_STOCKS):
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    sheetnames = get_sheetnames(spreadsheet_id)
    data= []
    for sheetname in sheetnames:
        if sheetname == 'Зеркала':
            continue
        result = sheet.values().get(spreadsheetId=spreadsheet_id,
                            range=f'{sheetname}!A:BS', majorDimension='ROWS').execute()
        values = result.get('values', [])
        
        for row in values:
            try:
                if len(row[0]) > 3 and row[0].isdigit():
                    data.append((row[0], row[1], row[5]))
            except:
                pass
    return data

def update_table_with_sum(source_id=SPREADSHEET_ID_STOCKS, target_id=SPREADSHEET_ID):
    data = get_all_data_from_google_sheet(source_id)
    i = 2
    body_data = []
    for barcode, article, stock in data:
        body_data.append([
            {'range': f'Баркоды!A{i}', 'values': [[barcode]]},
            {'range': f'Баркоды!B{i}', 'values': [[article]]},
            {'range': f'Баркоды!C{i}', 'values': [[stock]]},
            ])
        i+=1
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    body = {
        'valueInputOption': 'USER_ENTERED',
        'data': body_data}
    sheet.values().batchUpdate(spreadsheetId=target_id, body=body).execute()




if __name__ == '__main__':
    # get_all_data_from_google_sheet()
    update_table_with_sum()
    # service = build('sheets', 'v4', credentials=credentials)
    # sheet = service.spreadsheets()
    # sheet_metadata = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()
    # p
    # insert_data_in_table(get_counts_from_table(), SPREADSHEET_ID,FIRST_INDEX, spreadsheet_id=SPREADSHEET_ID)
    # insert_data_in_table
