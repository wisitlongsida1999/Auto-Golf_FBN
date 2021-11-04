import configparser
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select #For Dropdown
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep
import openpyxl
import pandas as pd
# config
my_config_parser = configparser.ConfigParser()
my_config_parser.read('config.ini')
payload = {
    'username': my_config_parser.get('GOLF_LOGIN','username'),
    'password': my_config_parser.get('GOLF_LOGIN','password'),
    'url':my_config_parser.get('GOLF_LOGIN','url'),
    'path':my_config_parser.get('GOLF_LOGIN','path')

    }

username = payload['username']
password = payload['password']
url = payload['url']
path = payload['path']


# driver=webdriver.Chrome()
# driver.get(url)
# sleep(1)
# url_before = driver.current_url
# WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="text"]'))).send_keys(username)
# sleep(1)
# WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="password"]'))).send_keys(password)
# sleep(1)
# WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="submit"]'))).click()
# sleep(1)

# while(url_before==driver.current_url):
#     pass

# driver.get(url)


df = pd.read_excel(path)
# print(df)
index = df.index
number_of_rows = len(index)
# print(number_of_rows)

# print(df['FA Report Acceptance'][0])
# print(df['Cisco Comment to supplier'][0])
# print(df['GOLF '][0])
dict_data= {}

for i in range(number_of_rows):
    dict_data[i] = [df['FA Report Acceptance'][i],df['Cisco Comment to supplier'][i],df['GOLF '][i]]

# print(dict_data)
print("="*100)

