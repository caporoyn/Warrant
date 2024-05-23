import time
import sqlite3
import pandas as pd
import json
import csv
import openpyxl
import os
import subprocess
import pyautogui


def open_and_save_with_numbers(input_file, input_file2):
    
    # 檢查文件OutXLS.xlsx是否存在
    if not os.path.exists(input_file):
        print(f"Error: 文件 '{input_file}' 不存在.")
        return
    if os.path.exists(input_file2):
        os.remove(input_file2)


    # 使用 subprocess 啟動 numbers 來打開 OutXLS.xlsx 並儲存為OutXLS.numbers
    try:
        subprocess.run(['open', '-a', 'Numbers', input_file])
        time.sleep(3)
        pyautogui.click()
        pyautogui.hotkey('command', 's')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.hotkey('command', 'w')
        time.sleep(2)

    except Exception as e:
        print(f"Error: 打开文件 '{input_file}' 失败.")
        print(e)


def export_numbers_to_excel(input_file, output_file):
    # 將 OutXLS.numbers 輸出為Excel

    if os.path.exists(output_file):
        os.remove(output_file)
    # 打開OutXLS.numbers後以pyautogui控制工具列匯出excel檔案
    subprocess.run(['open', '-a', 'Numbers', input_file])
    time.sleep(1)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'f2', interval = 0.2)
    pyautogui.press('right',2, interval = 0.2)
    pyautogui.press('down',11, interval=0.2)
    pyautogui.press('right', interval = 0.2)
    pyautogui.press('down', interval = 0.2)
    pyautogui.press('enter', interval = 0.2)
    pyautogui.press('enter', interval = 0.2)
    time.sleep(3)
    pyautogui.typewrite('OutXLS_processed', interval = 0.1)
    time.sleep(3)
    pyautogui.press('enter')
    pyautogui.hotkey('command', 'w')
    

if __name__ == "__main__":
    input_file = 'OutXLS.xlsx'
    input_file2 = 'OutXLS.numbers'
    output_file = 'OutXLS_processed.xlsx'
    
    open_and_save_with_numbers(input_file,input_file2)
    time.sleep(2)
    export_numbers_to_excel(input_file2, output_file)


