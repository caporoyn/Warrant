import time
import sqlite3
import pandas as pd
import json
import csv
import openpyxl
import sys
import json
import os

db = sqlite3.connect('warrants_db.db') 
cursor = db.cursor()

def run_query(query):
    return pd.read_sql_query(query, db)



# Query = '''
# SELECT "權證名稱", "實質槓桿" FROM WARS
# WHERE "實質槓桿" > 5;
# '''

Query = '''
SELECT "權證名稱", "實質<BR>槓桿", "成交價" / "Theta" * -1 AS cp_times FROM WARS
WHERE "實質<BR>槓桿" > 5 AND cp_times > 100
ORDER BY cp_times DESC ;
'''

# Query = '''
# SELECT "權證名稱", "實質<BR>槓桿", -1 * ("成交價" / "Theta") FROM WARS
# WHERE "實質<BR>槓桿" > 5 AND ;
# '''


result = run_query(Query)
print(result)