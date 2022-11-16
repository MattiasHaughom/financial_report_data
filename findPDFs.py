#%% Imports
import numpy as np
from datetime import timedelta
import datetime as datetime
from datetime import date
import pandas as pd
import os

directory = "H:\Personlig\Scripts"
os.chdir(directory)

finCalendar = pd.read_csv('finCal.csv')
finCalendar["Date"] = pd.to_datetime(finCalendar["Date"], infer_datetime_format= True)

finCalendar['event']

allTick = pd.read_excel('all_tickers.xlsx')

companies = allTick['aksje']


now = date.today()
today_timestamp = pd.to_datetime(now)
year = now.strftime("%Y")


reports_today = finCalendar.loc[finCalendar['Date'] == today_timestamp,:]
reports_today.reset_index(drop = True, inplace =True)



#%% SELENIUM navigate to the relevant website
# First part
from selenium.common.exceptions import TimeoutException
import ait
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


for number,comp in enumerate(todays_companies):

    url = "https://www.google.com/"
    
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    
    driver = webdriver.Chrome(options=options, service=Service(ChromeDriverManager().install()))
    driver.get(url)
    driver.maximize_window()
    
    
    button_id = 'W0wltc'
    
    button = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.ID,button_id)))
    driver.execute_script('arguments[0].click()', button)
    
    
    FinalXPath = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input'
    
    button = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH,FinalXPath)))
    driver.execute_script('arguments[0].click()', button)
    
    
    
    
    input_Field = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')))
    
    
    input_Field.send_keys('{c} {r}'.format(c = comp,r = reportSearch))
    
    
    
    
    FinalXPath = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[4]/center/input[1]'
    
    try:
        button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH,FinalXPath)))
    except TimeoutException:
        FinalXPath = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[2]/div[5]/center/input[1]'
        button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH,FinalXPath)))
        
    
    driver.execute_script('arguments[0].click()', button)
    
    
    sleep(1)
    ait.click(331,336)
    ait.click(331,336)
    sleep(2)
    
    investor_url = driver.current_url
    
    driver.close()
    driver.quit()
    

    #%% Get the latest report on the website
    reqs = requests.get(investor_url)
    soup = BeautifulSoup(reqs.text, 'html.parser')
    
    res = []
    for link in soup.find_all('a'):
        res.append(link.get('href'))
        
    urls = []
    for val in res:
        if val != None:
            urls.append(val)
        
    report_type = reports_today.loc[reports_today.index == number, 'event'].values[0]
    
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
    
    
    if len(report_type) > 1:
        l1 = [k for k in urls if year in k.lower() and any([j in k.lower() for j in report_type]) and any([y in k.lower() for y in report_wording])]
    elif len(report_type) == 1:
        l1 = [k for k in urls if year in k.lower() and report_type in k.lower() and any([y in k.lower() for y in report_wording])]
    
    not_available = []
    if len(l1) == 1:
        url = l1[0]
    elif len(l1) > 1:
        print("please check which report is relevant")
        url = l1[0]
    elif l1 == 0:
        not_available.append(comp)
        
    
    # URL of the image to be downloaded is defined as image_url
    directory = r"C:\Users\pj20\downloads"
    os.chdir(directory)
    
    headers = {
        'User-Agent' : 'PostmanRuntime/7.20.1',
        'Accept-Language': 'en-GB,en;q=0.5',
        'Link':investor_url,
        'Accept-Encoding': 'gzip, deflate',
        'Connection': 'keep-alive',
        'cache-control': 'no-cache',
        'Content-Type': 'text/html; charset=UTF-8'
        }
    
    # send a HTTP request to the server and save
    # the HTTP response in a response object called r
    r = requests.get(url) # stream=True, allow_redirects=True,
    r.headers
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
