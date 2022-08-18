import os.path
import sqlite3
import pandas as pd
import msoffcrypto
import io

def query_stockcode(stockcode):
    connection = sqlite3.connect(":memory:")
    cursor = connection.cursor()

    datacsv = pd.read_csv("MasterPartList.csv")
    datacsv.to_sql('masterpartlist', connection, if_exists='replace', index=False)

    query = f"""SELECT * FROM masterpartlist WHERE ItemCode LIKE '{stockcode}'"""

    with connection:
        cursor.execute(query)
        result = cursor.fetchall()

    return result


def create_db():
    d_excel = io.BytesIO()
    server_file = r"\\192.168.0.118\Engineering\MASTER PART LIST\Master Part List - CITEC.xlsx"
    if os.path.exists(server_file):
        with open(server_file, 'rb') as f:
            excelf = msoffcrypto.OfficeFile(f)
            excelf.load_key('CITEC')
            excelf.decrypt(d_excel)
        data_xls = pd.read_excel(d_excel, 'Sheet1', dtype=str)
        data_xls.to_csv('MasterPartList.csv', encoding='utf-8', header=True)
    else:
        create_offline_db()

def create_offline_db():
    excel = "C:\\Users\\chunh\\OneDrive\\OD - Documents\\Master Part List - CITEC.xlsx"
    data_xls = pd.read_excel(excel, 'MasterPartList', dtype=str)
    data_xls.to_csv('MasterPartList.csv', encoding='utf-8', header=True)