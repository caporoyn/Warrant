from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import sqlite3
import json
import sys

Path = "../Downloads/chromedriver-mac-arm64/chromedriver.exe"
driver = webdriver.Chrome()

class Input_data:
    def __init__(self, stock_id = '2330', warType = '全部', issuers = '全部',
               strikePrice_L = '', strikePrice_R = '', bsRate = '全部',
               inOut_L = '價內5%', inOut_R = '價外20%', esg = '',
               ivL = '40', ivR = '60', convertRate = '全部', lever = '全部',
               vol = '全部', period_L = '90日', period_R = '不選', outRate = '全部'):
        
        self.stock_id = stock_id
        self.warType = warType    
        self.issuers = issuers
        self.strikePrice_L = strikePrice_L
        self.strikePrice_R = strikePrice_R
        self.bsRate = bsRate
        self.inOut_L = inOut_L
        self.inOut_R = inOut_R
        self.esg = esg
        self.ivL = ivL
        self.ivR = ivR
        self.convertRate = convertRate
        self.lever = lever
        self.vol = vol
        self.period_L = period_L
        self.period_R = period_R
        self.outRate = outRate

# stock_id = sys.argv[1]
stock_id = input("請輸入股票代號：")
driver.get("https://warrant.kgi.com/edwebsite/views/warrantsearch/warrantsearch.aspx")
time.sleep(1)

data = Input_data()
data.stock_id = stock_id
# 標的股票
target = driver.find_element(By.CSS_SELECTOR, 'input[aria-owns="cmbUnderlying_listbox"]')
target.click()
target.send_keys(Keys.BACKSPACE)
target.send_keys(data.stock_id)
time.sleep(1)
target.send_keys(Keys.ENTER)

# 認購認售
time.sleep(2)
warType = driver.find_element(By.CSS_SELECTOR, 'span[aria-owns="ddlCP_listbox"]')
warType.click()
warType2 = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div#ddlCP-list > ul#ddlCP_listbox > li:nth-child(2)')))
warType2.click()

# 發行商
issuers = driver.find_element(By.CSS_SELECTOR, 'span[aria-owns="ddlIssuer_listbox"]')
issuers.click()
issuers2 = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div#ddlIssuer-list > ul#ddlIssuer_listbox > li:nth-child(1)')))
issuers2.click()

# # 到期日
period_L = driver.find_element(By.CSS_SELECTOR, 'span[aria-owns="numLastDaysFrom_listbox"]')
period_L.click()
period_L2 = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'div#numLastDaysFrom-list > ul#numLastDaysFrom_listbox > li:nth-child(4)')))
period_L2.click()

period_R = driver.find_element(By.CSS_SELECTOR, 'span[aria-owns="numLastDaysTo_listbox"]')
period_R.click()

period_R2 = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div#numLastDaysTo-list > ul#numLastDaysTo_listbox > li:nth-child(1)')))
period_R2.click()

# 買賣價差比
bsRate = driver.find_element(By.CSS_SELECTOR, 'span[aria-owns="ddlBidAskSpreadPercent_listbox"]')
bsRate.click()

bsRate2 = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'div#ddlBidAskSpreadPercent-list > ul#ddlBidAskSpreadPercent_listbox > li:nth-child(1)')))
bsRate2.click()

# 隱波
# ivL = driver.find_element(By.CSS_SELECTOR,'input#numImpVol')
ivR = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/div[3]/table[1]/tbody/tr[2]/td[6]/span/span/input[1]')
# ivR = driver.find_element(By.CSS_SELECTOR,'input#numImpVol')
# ivL.send_keys(data.ivL)
ivR.send_keys(data.ivR)

# 槓桿
lever = driver.find_element(By.CSS_SELECTOR, 'span[aria-owns="ddlLeverage_listbox"]')
lever.click()
lever2 = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'div#ddlLeverage-list > ul#ddlLeverage_listbox > li:nth-child(1)')))
lever2.click()

#價內外程度
inOut_L = driver.find_element(By.CSS_SELECTOR, 'span[aria-owns="ddlInOutPercentFrom_listbox"]')
inOut_L.click()

inOut_L2 = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'div#ddlInOutPercentFrom-list > ul#ddlInOutPercentFrom_listbox > li:nth-child(3)')))
inOut_L2.click()

inOut_R = driver.find_element(By.CSS_SELECTOR, 'span[aria-owns="ddlInOutPercentTo_listbox"]')
inOut_R.click()

inOut_R2 = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'div#ddlInOutPercentTo-list > ul#ddlInOutPercentTo_listbox > li:nth-child(8)')))
inOut_R2.click()

print("key in ok")
# 按下搜尋
opt_search = driver.find_element(By.CSS_SELECTOR, 'input[id = "btnQuery"]')
opt_search.click()


print("press ok")
# 下載excel
time.sleep(5)
make_excel = driver.find_element(By.CSS_SELECTOR, 'input[type = "button"][onclick *= "saveAsExcel"]')
make_excel.click()

print("download ok")

time.sleep(10)


driver.quit()