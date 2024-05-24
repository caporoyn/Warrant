使用指南：
step1:
執行"python3 warrant_search_kgi.py"，並依指示輸入股票代號，由selenium自動化操作凱基權證網下載資料

step2:
///使用前先將檔案中之路徑調整為自身之路徑///
執行"python3 excel_process.py"將下載檔案"OutXLS.xlsx"轉換為"OutXLS_pro.xlsx"，此步驟以pyautogui完成，解決mac開啟xlsx引發之格式問題


step3:
執行"python3 make_database.py"將"OutXLS_pro.xlsx"匯入以sqlite3建立之資料庫中

step4:
執行"query.py"篩選出合適之標的，若會使用sql語法也可自行變更篩選策略 
