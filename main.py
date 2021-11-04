import configparser
from selenium import webdriver


# config
my_config_parser = configparser.ConfigParser()
my_config_parser.read('config.ini')
payload = {
    'username': my_config_parser.get('GOLF_LOGIN','username'),
    'password': my_config_parser.get('GOLF_LOGIN','password')
    }

target_web = ''

driver=webdriver.Chrome()
driver.get(target_web)

