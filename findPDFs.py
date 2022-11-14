#%% Imports
import numpy as np
import pandas as pd
import os
from datetime import timedelta
import datetime as datetime

directory = "H:\Personlig\Scripts"
os.chdir(directory)

finCal = pd.read_csv('fincal.csv')
allTick = pd.read_excel('all_tickers.xlsx')


#%% SELENIUM get all financial statement links
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import ait
import requests


companies = allTick['aksje']


cmp = companies[0]
report =  "financial reports"


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


input_Field.send_keys('{c} {r}'.format(c = cmp,r = report))




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



url = "https://2020bulkers.com/content/uploads/2022/11/2020-Bulkers-Consolidated-financial-statements-Q3-2022.pdf"
  
# URL of the image to be downloaded is defined as image_url
r = requests.get(url)

# send a HTTP request to the server and save
# the HTTP response in a response object called r
with open("2020-Bulkers-Consolidated-financial-statements-Q3-2022.pdf",'wb') as f:
  
    # Saving received content as a png file in
    # binary format
  
    # write the contents of the response (r.content)
    # to a new file in binary mode.
    f.write(r.content)


