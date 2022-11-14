
#%% SELENIUM get table
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import os



url = "https://live.euronext.com/en/listview/financial-events"


options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])

driver = webdriver.Chrome(options=options, service=Service(ChromeDriverManager().install()))
driver.get(url)


sleep(5)


first = []
second = []
third = []


lst = [2,5,6]+[7]*36+[8,9]
table = range(0,41)
for page,i in zip(lst,table):
    # FOR OTHER DATA POINTS
    before_XPath = '//*[@id="DataTables_Table_{}"]/tbody/tr['.format(i)
    aftertd_XPath_1 = ']/td[1]/time'
    aftertd_XPath_2 = ']/td[2]/span'
    aftertd_XPath_3 = ']/td[3]/a'

    sleep(3)
     
    rows = len(driver.find_elements(By.XPATH,'//*[@id="DataTables_Table_{}"]/tbody/tr'.format(i)))
    
    for t_row in range(1,(rows+1),2):
        FinalXPath = before_XPath + str(t_row) + aftertd_XPath_1
        elem = driver.find_element(By.XPATH,FinalXPath)
        first.append(elem.text)
     
    for t_row in range(1,(rows+1),2):
        FinalXPath = before_XPath + str(t_row) + aftertd_XPath_2
        elem = driver.find_element(By.XPATH,FinalXPath)
        second.append(elem.text)
    
    for t_row in range(1,(rows+1),2):
        FinalXPath = before_XPath + str(t_row) + aftertd_XPath_3
        try:
            elem = driver.find_element(By.XPATH,FinalXPath)
        except NoSuchElementException:
            elem = driver.find_element(By.XPATH,FinalXPath[:-2])
        third.append(elem.text)
    
    
    before_XPath = '//*[@id="content"]/section/div[2]/div/div[4]/div/nav/ul/li['
    aftertd_XPath_1 = ']/a'
    FinalXPath = before_XPath + str(page) + aftertd_XPath_1
    
    button = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH,FinalXPath)))
    driver.execute_script('arguments[0].click()', button)
            


tmp = {'Date':first, 'company': second,'event': third}
finCal = pd.DataFrame(data = tmp)
directory = "H:\Personlig\Scripts"
os.chdir(directory)
allTick = pd.read_excel('all_tickers.xlsx')

companies = allTick['aksje']


finCalendar = []
for number,cmp in enumerate(finCal['company']):
    if cmp in list(companies):
        finCalendar.append(finCal.loc[finCal.index == number,:])
finCalendar = pd.concat(finCalendar)

finCalendar.to_csv('finCal.csv', index = False)
