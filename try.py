import configparser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep
import pandas as pd


my_config_parser = configparser.ConfigParser()
my_config_parser.read('config.ini')
payload = {
    'username': my_config_parser.get('GOLF_LOGIN','username'),
    'password': my_config_parser.get('GOLF_LOGIN','password'),
    'url':my_config_parser.get('GOLF_LOGIN','url'),
    'path':my_config_parser.get('GOLF_LOGIN','path')

    }
global dict_all_span
global dict_excel_data
global path
global url
global dict_state_incorrect
global dict_closure_incorrect


username = payload['username']
password = payload['password']
url = payload['url']
path = payload['path']
dict_all_span = {} #global
dict_excel_data = {} #global
time_out = 30 #default
dict_state_incorrect = {}
dict_closure_incorrect = {}


#Set Up
global driver
driver=webdriver.Chrome()
driver.minimize_window()
driver.get(url)
url_before = driver.current_url
WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="text"]'))).send_keys(username)

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="password"]'))).send_keys(password)

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="submit"]'))).click()


time = 0
while(url_before==driver.current_url):
    sleep(1)
    time+=1
    print("Waiting for response from server",time_out,"seconds")
    if(time > time_out):
        raise Exception("Fail Connected to server")


driver.get("view-source:https://golf.fabrinet.co.th/normaluser/WorkFlow.asp?rnt=3383386-CISCO&mode=1&documentgroup=&documentprefix=All%20Documents&docstatus=&orderby=&CurPage=1&uid=15!52!38!575&searchid=3383386-CISCO")
print('='*100)
driver.page_source()
WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="button"][@name="viewi21354816s17c2924t6"]'))).click()