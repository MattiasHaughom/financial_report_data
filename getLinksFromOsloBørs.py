#%% Imports
import numpy as np
from datetime import timedelta
import datetime as datetime
from datetime import date
import pandas as pd
import os

directory = "H:\Personlig\Scripts"
os.chdir(directory)

finCalendar = pd.read_csv('finCalv2.csv')
finCalendar["Date"] = pd.to_datetime(finCalendar["Date"], infer_datetime_format= True)


allTick = pd.read_excel('all_tickers.xlsx')
companies = allTick['aksje']
ticker = allTick['ticker']


now = date.today()
today_timestamp = pd.to_datetime(now)
year = now.strftime("%Y")


reports_today = finCalendar.loc[finCalendar['Date'] == today_timestamp,:]
reports_today.reset_index(drop = True, inplace =True)

result = []
for number,cmp in enumerate(reports_today['company']):
    if cmp in list(companies):
        result.append(finCalendar.loc[finCalendar.index == number,:])
result = pd.concat(result)

result.to_csv('finCal.csv', index = False)

#%% SELENIUM navigate to the relevant website
# First part
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
# Second part
import requests
from bs4 import BeautifulSoup
from pathlib import Path
# Third part
import PyPDF2
import tabula
from tabula import read_pdf
import re


todays_companies = reports_today['company']
reportSearch =  "financial reports"


url = "https://newsweb.oslobors.no/"

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.headless = True

driver = webdriver.Chrome(options=options, service=Service(ChromeDriverManager().install()))

driver.get(url)

WebDriverWait(driver, 15).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/main/table/tbody/tr')))

rows = len(driver.find_elements(By.XPATH,'//*[@id="root"]/div/main/table/tbody/tr'))
before_XPath = '//*[@id="root"]/div/main/table/tbody/tr['
aftertd_XPath1 = ']/td[4]/a'
aftertd_XPath2 = ']/td[3]'
aftertd_XPath3 = ']/td[6]'

links = []
tekst = []
for t_row in range(1, (rows + 1)):
    FinalXPath = before_XPath + str(t_row) + aftertd_XPath1
    links.append(driver.find_element(By.XPATH,FinalXPath).get_attribute("href"))
    tekst.append(driver.find_element(By.XPATH,FinalXPath).text)

bors_ticker = []
for t_row in range(1, (rows + 1)):
    FinalXPath = before_XPath + str(t_row) + aftertd_XPath2
    bors_ticker.append(driver.find_element(By.XPATH,FinalXPath).text)

number_of_files = []
for t_row in range(1, (rows + 1)):
    FinalXPath = before_XPath + str(t_row) + aftertd_XPath3
    bors_ticker.append(driver.find_element(By.XPATH,FinalXPath).text)



d = {"links":links,"tekst":tekst, "ticker":bors_ticker, "files": number_of_files}
data = pd.DataFrame(data = d)

investor_url = data.loc[["AURORA EIENDOM AS".lower() in j.lower() for j in data['ticker']],'links'].values[0]


#%% Get the latest report on the website

driver.get(investor_url)

WebDriverWait(driver, 15).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/main/div[2]/div[2]/div[1]/div[9]/ul/li')))


files = len(driver.find_elements(By.XPATH,'//*[@id="root"]/div/main/div[2]/div[2]/div[1]/div[9]/ul'))

                                 
# //*[@id="root"]/div/main/div[2]/div[2]/div[1]/div[8]/ul/li
# //*[@id="root"]/div/main/div[2]/div[2]/div[1]/div[9]/ul/li
## if one
# //*[@id="root"]/div/main/div[2]/div[2]/div[1]/div[9]/ul/li/a
## if two
# //*[@id="root"]/div/main/div[2]/div[2]/div[1]/div[9]/ul/li[1]/a
# //*[@id="root"]/div/main/div[2]/div[2]/div[1]/div[9]/ul/li[2]/a




# URL of the image to be downloaded is defined as image_url
directory = r"C:\Users\pj20\downloads"
os.chdir(directory)



# send a HTTP request to the server and save
# the HTTP response in a response object called r
r = requests.get(url, headers= headers) # stream=True, allow_redirects=True,
filename = Path(l1[0].split("/")[-1].strip())
filename.write_bytes(r.content)


#%% check the content of the PDF

report = max(os.listdir(directory), key = os.path.getctime)

doc = PyPDF2.PdfFileReader(report, strict=False)

pages = doc.getNumPages()

# Seach for terms in the document
incomeStmPages = []
incomeCompPages= []
opProfitPages= []
netProfitPages= []
profitPeriodPages= []
profitPages = []

income_statement = 'income statement'
comp_income = 'comprehensive income'
op_profit = 'operating profit'
net_profit = 'net profit'
profit_period = 'profit for the period'
profit = 'profit'

for i in range(pages):
    current_page = doc.getPage(i)
    text = current_page.extractText()
    if re.findall(income_statement, text.lower()):
        incomeStmPages.append(i)
    if re.findall(comp_income, text.lower()):
        incomeCompPages.append(i)
    if re.findall(op_profit, text.lower()):
        opProfitPages.append(i)
    if re.findall(net_profit, text.lower()):
        netProfitPages.append(i)
    if re.findall(profit_period, text.lower()):
        profitPeriodPages.append(i)
    if re.findall(profit, text.lower()):
        profitPages.append(i)


if len(incomeStmPages) == 1:
    result = incomeStmPages[0]
    pass
elif (len(incomeStmPages) > 1 and (len(opProfitPages) > 1 or len(opProfitPages) > 1 or len(opProfitPages) > 1) and len(profitPages>1)):
    p = [opProfitPages,netProfitPages,profitPages]
    tmp = []
    for i in p:
        tmp.append([x for x in incomeStmPages if x in i])
    c1set = frozenset(tmp[0])
    c2 = tmp[1:]
    result = [n for lst in c2 for n in lst if n in c1set]
    # check comprehensive income instead
elif len(incomeStmPages) == 0 and len(incomeCompPages) > 0:
    p = [opProfitPages,netProfitPages,profitPages]
    tmp = []
    for i in p:
        tmp.append([x for x in incomeCompPages if x in i])
    c1set = frozenset(tmp[0])
    c2 = tmp[1:]
    result = [n for lst in c2 for n in lst if n in c1set]


if len(result) == 1:
    result = [result[0] +1]
elif len(result) > 1:
    result = [i + 1 for i in result]


if len(result) > 1:
    tables = tabula.read_pdf(report,pages=result,multiple_tables=True,guess=True, stream=True)
    table = {}
    for number,i in enumerate(tables):
        table[number] = i
else:
    tables = tabula.read_pdf(report,pages=result[0],multiple_tables=True,guess=True, stream=True)
    table = pd.concat(tables)

result_Income = {}
result_Income[comp] = table
