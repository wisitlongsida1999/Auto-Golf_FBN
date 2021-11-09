import configparser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep
import pandas as pd




def initialize():
    # config
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

    read_excel_data()

    time = 0
    while(url_before==driver.current_url):
        sleep(1)
        time+=1
        print("Waiting for response from server",time_out,"seconds")
        if(time > time_out):
            raise Exception("Fail Connected to server")


def read_excel_data():
    df = pd.read_excel(path)
    print(df)
    index = df.index
    number_of_rows = len(index)
    print(number_of_rows)

    print(df['FA Report Acceptance'][0])
    print(df['Cisco Comment to supplier'][0])
    print(df['GOLF '][0])

    for i in range(number_of_rows):
        dict_excel_data[df['GOLF '][i]] = [df['FA Report Acceptance'][i],df['Cisco Comment to supplier'][i]]

    print(dict_excel_data)
    print("="*100)




def fill_in_form(cisco_id):

    frame = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//frame[@name="down"]')))

    driver.switch_to_frame(frame)

    WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, '//textarea[@name="i0s2c211"]'))).send_keys(dict_excel_data[cisco_id][1])


    if 'final' in dict_excel_data[cisco_id][0].lower() and 'replacement' in dict_excel_data[cisco_id][1].lower():

        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, '//input[@type="radio"][@name="i0s4c50t14"][@value="129"]'))).click()
    
    elif 'final' in dict_excel_data[cisco_id][0].lower() and 'scrap' in dict_excel_data[cisco_id][1].lower():

        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, '//input[@type="radio"][@name="i0s4c50t14"][@value="114"]'))).click()

    elif 'recall' in dict_excel_data[cisco_id][1].lower():

        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, '//input[@type="radio"][@name="i0s4c50t14"][@value="130"]'))).click()

    elif 'reject' in dict_excel_data[cisco_id][0].lower() or 'prelim' in dict_excel_data[cisco_id][0].lower():

        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, '//input[@type="radio"][@name="i0s4c50t14"][@value="23"]'))).click()

    else:

        dict_closure_incorrect[cisco_id] = dict_excel_data[cisco_id]


    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@type="button"][@name="hello"][@value="Send  Your  Action"]'))).click()
    

    # handle with alert
    try:
        WebDriverWait(driver, 3).until(ec.alert_is_present())

        alert = driver.switch_to.alert
        alert.accept()
        print("alert accepted")

    except:
        print("no alert")

#Access to the golf id


def access_to_golf_id(cisco_id):
    driver.get(url)
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="searchid"]'))).send_keys(cisco_id)
    WebDriverWait(driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//input[@name="searchbutton"]')))[0].click()
    print('='*200)
    all_span = WebDriverWait(driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//span[@style="cursor:hand"]')))


    print('='*200)
    all_span_text = []
    for i in all_span:
        print(i.text)
        all_span_text.append(i.text)

    dict_all_span[all_span_text[3]] = all_span_text
    print('='*200)

    print("Number of spans are",len(all_span),"Expected to \"12\" ")
    all_span[9].text
    if all_span_text[3] == cisco_id and 'FA: Review' in all_span_text[10]:
        print('='*200)

        print(all_span_text[3],"is \"Review State\" ")
        all_span[3].click()
        fill_in_form(cisco_id)
    else:
        print(all_span_text[3],"is not \"Review State\" ")
        dict_state_incorrect[cisco_id] = dict_excel_data[cisco_id]






def main():

    initialize()
    sleep(1)
    
    if __name__ == '__main__':

        for cisco_id in dict_excel_data:
            access_to_golf_id(cisco_id)
            sleep(1)
            
        
        print(  """
                ─────────███──────────███
                ────────█▓▓▓█────────█▓▓▓█
                ───────█▓▓▓▓██████████▓▓▓▓█
                ───────█▓▓▓██▒▒▒▒▒▒▒▒██▓▓▓█
                ────────████▒▒▒▒▒▒▒▒▒▒████
                ──────────█▒▒▒██▒▒██▒▒▒█
                ──────────█▒▒▒▒▒▒▒▒▒▒▒▒█
                ──────────█▒▒▒▒░██░▒▒▒▒█
                ──────────██▒▒▓▒▒▒▒▓▒▒██
                ───────────██▒▒▓▓▓▓▒▒██
                ────────────█▒▒▒▒▒▒▒▒█
                ─────████████████████████████
                ───▓▓▓▓▒▒░░░░░░░░░░░░░░░░▒▒▓▓▓▓
                ──▓▓▒▒▓▒░░░░░░░░░░░░░░░░░░▒▓▒▒▓▓
                ──▓▒▒▒▓░░░ *♥Finish!*♥  ░░▒▓▒▒▒▓
                ──▓▓▒▒▓▒▒░░░░░░░░░░░░░░░░▒▒▓▒▒▓▓
                ───▓▓▒▓▒▒░░░░░░░░░░░░░░░░▒▒▓▒▓▓ 
                ─────████████████████████████
                ─────────█▒▒▓▓▓▓▓▓▓▓▓▒▒█
                ────────██▒▒▓▓▓▓▓▓▓▓▓▒▒██
                ───────-█▓▒▒▒▓▓▓▓▓▓▓▓▓▒▒▒▓█
                ─███▄─█▒▒▓▒▒▒▓▓▓▓▓▓▓▒▒▒▓▒▒█─▄███
                █▓▓▓██▒▒▒▓▒▒▒▒▒▒▒▒▒▒▒▒▒▓▒▒▒██▓▓▓█
                █▓▓▓▓█▒▒▒▒▓▒▒▒▒▒▒▒▒▒▒▒▓▒▒▒▒█▓▓▓▓█
                █▓▓▓▓██▒▒▒▒▓█████████▓▒▒▒▒██▓▓▓▓█
                █▓▓▓▓▓█▒▒▒█▓─────────▓█▒▒▒█▓▓▓▓▓█
                ─█▓▓▓▓█▒██▓───────────▓██▒█▓▓▓▓█
                ──█▓▓▓██▓▓─────────────▓███▓▓▓█
                ───████▓────────────────▓█████

            """     )

        if dict_closure_incorrect.__len__() != 0 :
            print('='*100)
            print('Show closure incorrect')
            print('-'*100)
            print(dict_closure_incorrect)
        
        if dict_state_incorrect.__len__() != 0:
            print('='*100)
            print('Show state incorrect')
            print('-'*100)
            print(dict_state_incorrect)

        while(input("Press \'e\' to exit !\n").lower() != 'e'):

            pass

        print("Exit the program")
        sleep(2)
        driver.quit()
        


main()
