#%% Imports
import numpy as np
from datetime import timedelta
import datetime as datetime
from datetime import date
import pandas as pd
import os

directory = "H:\Personlig\Scripts"
os.chdir(directory)

# Define date and time variables 
now = date.today()
today_timestamp = pd.to_datetime(now)
year = now.strftime("%Y")

# Download data
finCalendar = pd.read_csv('finCalv2.csv')
finCalendar["Date"] = pd.to_datetime(finCalendar["Date"], infer_datetime_format= True)
allTick = pd.read_excel('all_tickers.xlsx')
companies = allTick['aksje']

# Identify the companies releasing report today
reports_today = finCalendar.loc[finCalendar['Date'] == today_timestamp,:]
reports_today.reset_index(drop = True, inplace =True)
# remove company title
reports_today['company'] = reports_today['company'].str.replace("\s+\S+$", "")
lst = [i.lower() for i in reports_today['company']]
lst = '|'.join(lst)

companies = pd.Series([i.lower() for i in companies])
public_tickers = allTick.loc[companies.str.contains(lst),:]




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


todays_companies = public_tickers['ticker']


report_type = reports_today.loc[:, 'event'].unique()[0]

if "Q3" in report_type:
    report_type = ["q3","third","3st"]
elif "Q2" in report_type:
    report_type = ["q2","second"]
elif "Q1" in report_type:
    report_type = ["q1","first"]
elif "Q4" in report_type:
    report_type = ["q4","fourth"]
elif "Annual Report" in report_type:
    report_type = "annual report"
elif "Half-yearly" in report_type:
    report_type = ["semi annual","1H"]

report_wording = ["statement","report"]

words_to_check = "|".join(report_type+report_wording)


url = "https://newsweb.oslobors.no/"


options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
#options.headless = True

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
    number_of_files.append(driver.find_element(By.XPATH,FinalXPath).text)


d = {"links":links,"tekst":tekst, "ticker":bors_ticker, "files": number_of_files}
data = pd.DataFrame(data = d)
    
event_text = []
result_Income = {}
for tick in todays_companies:
    
    tmp = data.loc[[tick.lower() in j.lower() for j in data['ticker']],['links','tekst']] #!!!
    
    if len(tmp['links']) == 1:
        investor_url = tmp['links'].values[0]
    elif len(tmp['links']) > 1:
        tmp['tekst'] = [i.lower() for i in tmp['tekst']]
        investor_url = tmp.loc[tmp['tekst'].str.contains(words_to_check),'links'].values[0]


    #%% find the report element and press it to download the report
    
    driver.get(investor_url)
    
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/main/div[2]')))
    
    
    number_of_links = data.loc[[tick.lower() in j.lower() for j in data['ticker']],'files'].values[0] #!!!
    

    all_links =  driver.find_elements(By.TAG_NAME,'a')
    
    if len(report_type) > 1:
        element = [k.text for k in all_links if year in k.text.lower() and any([j in k.text.lower() for j in report_type]) or any([y in k.text.lower() for y in report_wording])]
    elif len(report_type) == 1:
        element = [k.text for k in all_links if year in k.text.lower() and report_type in k.text.lower() or any([y in k.text.lower() for y in report_wording])]

    button = WebDriverWait(driver, 25).until(
        EC.element_to_be_clickable((By.LINK_TEXT,"{}".format(element[0]))))
        
    driver.execute_script('arguments[0].click()', button)
    sleep(2.5)
                
    try:
        event_text.append(driver.find_element(By.XPATH,'//*[@id="root"]/div/main/div[2]/div[2]/div[2]').text)
    except NoSuchElementException:
        pass
    
    
    #%% check the content of the PDF
    sleep(5)
    # URL of the image to be downloaded is defined as image_url
    directory = r"C:\Users\pj20\downloads"
    os.chdir(directory)
    
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
    consolidatedProfitPages = []
    profitandLossPages = []
    
    # statement
    income_statement = 'income statement'
    comp_income = 'comprehensive income'
    consolidated_profit = 'consolidated profit and loss'
    profit_and_loss = "profit & loss "

    
    # key words
    op_profit = 'operating profit'
    net_profit = 'net profit'
    profit_period = 'profit for the period'
    profit = 'profit'
    
    for i in range(pages):
        current_page = doc.getPage(i)
        text = current_page.extractText()
        
        #
        if re.findall(income_statement, text.lower()):
            incomeStmPages.append(i)
        if re.findall(comp_income, text.lower()):
            incomeCompPages.append(i)
        if re.findall(consolidated_profit, text.lower()):
            consolidatedProfitPages.append(i)
        if re.findall(profit_and_loss, text.lower()):
            profitandLossPages.append(i)    

    
        if re.findall(op_profit, text.lower()):
            opProfitPages.append(i)
        if re.findall(net_profit, text.lower()):
            netProfitPages.append(i)
        if re.findall(profit_period, text.lower()):
            profitPeriodPages.append(i)
        if re.findall(profit, text.lower()):
            profitPages.append(i)

    
    if len(incomeStmPages) == 1:
        result = [incomeStmPages[0]]
    elif (len(incomeStmPages) > 1 and (len(opProfitPages) > 1 or len(netProfitPages) > 1 or len(profitPeriodPages) > 1) and len(profitPages)>1):
        p = [opProfitPages,netProfitPages,profitPages]
        tmp = []
        for i in p:
            tmp.append([x for x in incomeStmPages if x in i])
        c1set = frozenset(tmp[0])
        c2 = tmp[1:]
        result = [n for lst in c2 for n in lst if n in c1set] 
    elif len(incomeStmPages) == 0 and any([len(type_s) > 0 for type_s in [incomeCompPages, consolidatedProfitPages, profitandLossPages]]):
        state = [incomeCompPages, consolidatedProfitPages, profitandLossPages]
        p = [opProfitPages,netProfitPages]
        tmp = []
        tmp.append([x for x in state if [x in j for j in p]])
        t = []
        for i in tmp:
            for j in i:
                print(j)
                if len(j) > len(t):
                    t = j
        result = t


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
    
    result_Income[tick] = table
    
driver.close()
driver.quit()