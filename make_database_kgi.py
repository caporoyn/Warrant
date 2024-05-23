import time
import sqlite3
import pandas as pd
import json
import csv
import openpyxl

column_names = ['權證代碼', '權證名稱','買價','買量','賣價','賣量',
'成交價', '漲跌幅%', '成交量', '買賣價<BR>差比%','BIV','SIV', '履約價',
 '行使<BR>比例', '價內外<br>(%)', '剩餘<BR>天數', '實質<BR>槓桿',
  'Delta', 'Theta', '流通<BR>在外(%)']

#必須先將下載項目中的OutXLLS.xlsx移動到此資料夾
excel = pd.read_excel('OutXLS_processed.xlsx')
data_list = excel.values.tolist()

# 使用sqlite3連線/創建資料庫
db = sqlite3.connect('warrants_db.db') 
cursor = db.cursor()
cursor.execute('''
DROP TABLE WARS;
''')

cursor.execute('''
    CREATE TABLE WARS
    (
        '權證代碼'          CHAR(10),
        '權證名稱'          CHAR(20),
        '買價'             FLOAT,
        '買量'             INT,
        '賣價'             FLOAT,
        '賣量'             INT,
        '成交價'           FLOAT,
        '漲跌幅%'          FLOAT,
        '成交量'           INT,   
        '買賣價<BR>差比%'   FLOAT,
        'BIV'             FLOAT,
        'SIV'             FLOAT,
        '履約價'           FLOAT,
        '行使<BR>比例'     FLOAT,
        '價內外<br>(%)'    FLOAT,
        '剩餘<BR>天數'      INT,
        '實質<BR>槓桿'      FLOAT,
        'Delta'            FLOAT,
        'Theta'             FLOAT,
        '流通<BR>在外(%)'     FLOAT
    );
''')
db.commit()

#將數據輸入資料庫
for data in data_list:
    cursor.execute("INSERT INTO WARS VALUES"+str(tuple(data)))

db.commit()

query = '''
SELECT "權證名稱", "實質<BR>槓桿" FROM WARS 
WHERE "成交量" > 50;
'''
result = pd.read_sql_query(query, db)
print(result)



