import configparser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep
import pandas as pd
import sys
import traceback







# config
my_config_parser = configparser.ConfigParser()
my_config_parser.read('config.ini')
payload = {
    'username': my_config_parser.get('GOLF_LOGIN','yanee_id'),
    'password': my_config_parser.get('GOLF_LOGIN','yanee_pwd'),

    }


username = payload['username']
password = payload['password']