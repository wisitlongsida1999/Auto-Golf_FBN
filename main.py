import configparser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep
import pandas as pd
import sys
import traceback




def initialize():

    global dict_all_span
    global dict_excel_data
    global dict_state_incorrect
    global dict_closure_incorrect
    global not_found_case
    global owner_golf
    global time_out

    dict_all_span = {} #global
    dict_excel_data = {} #global
    time_out = 30 #default
    dict_state_incorrect = {}
    dict_closure_incorrect = {}
    not_found_case = []
    owner_golf = {'nathawit_golf':{},'arissara_golf':{},'wisit_golf':{},'yanee_golf':{}}




def login(user,pwd):

    # config
    my_config_parser = configparser.ConfigParser()
    my_config_parser.read('config.ini')
    payload = {
        'username': my_config_parser.get('GOLF_LOGIN',f'{user}'),
        'password': my_config_parser.get('GOLF_LOGIN',f'{pwd}'),

        }


    username = payload['username']
    password = payload['password']




    global driver
    driver=webdriver.Chrome()
    # driver.minimize_window()
    driver.get('https://golf.fabrinet.co.th/normaluser/MyWorkFlow.asp?mode=1&documentgroup=&documentprefix=&docstatus=1')
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




def read_excel_data():
    df = pd.read_excel('GOLF_8D_Template.xlsm')
    print(df)
    index = df.index
    number_of_rows = len(index)
    print(number_of_rows)


    for i in range(number_of_rows):

        if '400' in df['PID'][i]:
            
            owner_golf['arissara_golf'][df['GOLF '][i]] = [df['FA Report Acceptance'][i],df['Cisco Comment to supplier'][i],df['PID'][i]]

        elif '40/100' in df['PID'][i] :

            owner_golf['nathawit_golf'][df['GOLF '][i]] = [df['FA Report Acceptance'][i],df['Cisco Comment to supplier'][i],df['PID'][i]]
        
        elif '40' in df['PID'][i] or '100' in df['PID'][i] or '4SFP10G' in df['PID'][i] or '4X10'in df['PID'][i]:
            
            owner_golf['wisit_golf'][df['GOLF '][i]] = [df['FA Report Acceptance'][i],df['Cisco Comment to supplier'][i],df['PID'][i]]

        else:

            owner_golf['yanee_golf'][df['GOLF '][i]] = [df['FA Report Acceptance'][i],df['Cisco Comment to supplier'][i],df['PID'][i]]

    print(owner_golf)
    print("="*100)


def fill_in_form(cisco_id,owner):

    frame = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//frame[@name="down"]')))

    driver.switch_to.frame(frame)

    WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, '//textarea[@name="i0s2c211"]'))).send_keys(owner_golf[owner][cisco_id][1] + ' ( PID: ' + owner_golf[owner][cisco_id][2]+' )')

    if ('final' in owner_golf[owner][cisco_id][0].lower() and 'replacement' in owner_golf[owner][cisco_id][1].lower() ) or ('final' in owner_golf[owner][cisco_id][0].lower() and 'credit' in owner_golf[owner][cisco_id][1].lower()):

        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, '//input[@type="radio"][@name="i0s3c50t14"][@value="129"]'))).click()

    elif 'recall' in owner_golf[owner][cisco_id][1].lower() or 're-call' in owner_golf[owner][cisco_id][1].lower():

        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, '//input[@type="radio"][@name="i0s3c50t14"][@value="130"]'))).click()
    
    elif 'final' in owner_golf[owner][cisco_id][0].lower() and 'scrap' in owner_golf[owner][cisco_id][1].lower():

        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, '//input[@type="radio"][@name="i0s3c50t14"][@value="114"]'))).click()

    elif ('reject' in owner_golf[owner][cisco_id][0].lower() or 'prelim' in owner_golf[owner][cisco_id][0].lower() or 'final' in owner_golf[owner][cisco_id][0].lower()) and ('replacement' not in owner_golf[owner][cisco_id][1].lower() and 'scrap' not in owner_golf[owner][cisco_id][1].lower() and 'recall' not in owner_golf[owner][cisco_id][1].lower() and 're-call' not in owner_golf[owner][cisco_id][1].lower()  ):

        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, '//input[@type="radio"][@name="i0s3c50t14"][@value="261"]'))).click()

    else:

        dict_closure_incorrect[cisco_id] = owner_golf[owner][cisco_id]

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


def access_to_golf_id(cisco_id,owner):
    driver.get('https://golf.fabrinet.co.th/normaluser/MyWorkFlow.asp?mode=1&documentgroup=&documentprefix=&docstatus=1')
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="searchid"]'))).send_keys(cisco_id)
    WebDriverWait(driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//input[@name="searchbutton"]')))[0].click()

    try:
        all_td = WebDriverWait(driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//tr[@height="25"]//td')))

        print('='*200)
        all_td_text = []
        for i in all_td:
            print(i.text)
            all_td_text.append(i.text)

        print('='*200)

        print("Number of spans are",len(all_td),"Expected to \"11\" ")

        if all_td_text[1].strip() == cisco_id and 'FA: Review' in all_td_text[8]:
            print('='*200)

            print(all_td_text[1],"is \"Review State\" ")
            all_td[1].click()
            fill_in_form(cisco_id,owner)
        else:
            print('*'*100)
            # print('<<<',all_td_text[1],'<<<',cisco_id)
            print(all_td_text[1],"is not \"Review State\" ")
            dict_state_incorrect[cisco_id] = owner_golf[owner][cisco_id]

    except:
        
        not_found_case.append(cisco_id)

        traceback.print_exc()






def main():


    
    if __name__ == '__main__':

        initialize()

        read_excel_data()

        for owner in owner_golf:

            if owner_golf[owner].__len__() > 0:

                user = owner.split('_')[0] + '_id'
                pwd = owner.split('_')[0] + '_pwd'

                login(user,pwd)

                sleep(1)

                for cisco_id in owner_golf[owner]:
                    access_to_golf_id(cisco_id,owner)
                    sleep(1)

                driver.quit()

                
            
        
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

        if not_found_case.__len__() != 0 :
            print('='*100)
            print('Show Not found Case')
            print('-'*100)
            print(not_found_case)

        while(input("Press \'e\' to exit !\n").lower() != 'e'):

            pass

        print("Exit the program")
        sleep(2)
        driver.quit()
        


main()
