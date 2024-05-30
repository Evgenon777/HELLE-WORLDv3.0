from openpyxl.utils import get_column_letter
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv()

KEY_TABLE = os.getenv('KEY_TABLE')

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name("secret.json", scope)

def CopyFromExcInGsh():
    client = gspread.authorize(credentials)

    spreadsheet = client.open_by_key(KEY_TABLE)
    worksheet = spreadsheet.worksheet('Аналитика и статистика все компании')

    df = pd.read_excel("sheet.xlsx")
    data_list = df.values.tolist()
    num_cols = len(data_list[0])

    cell_list = worksheet.range('A1:' + get_column_letter(num_cols) + str(len(data_list)))
    for cell in cell_list:
        row = (cell.row - 1) if (cell.row - 1) < len(data_list) else -1
        col = (cell.col - 1) if (cell.col - 1) < num_cols else -1
        if row != -1 and col != -1:
            value = data_list[row][col]
            if pd.notna(value):
                cell.value = str(value)

    worksheet.update_cells(cell_list)
    print("Данные успешно загружены в таблицу Google Sheets!")