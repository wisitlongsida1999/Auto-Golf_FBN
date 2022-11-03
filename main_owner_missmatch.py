import configparser
from selenium import webdriver
from selenium.webdriver.common.by import by
from selenium.webdriver.support.ui import webdriverwait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep
import pandas as pd
import sys
import traceback
import chromedriver_autoinstaller
from selenium.webdriver.chrome.service import service



def initialize():

    global dict_all_span
    global dict_excel_data
    global dict_state_incorrect
    global dict_closure_incorrect
    global not_found_case
    global owner_golf
    global time_out
    global driver_path

    dict_all_span = {} #global
    dict_excel_data = {} #global
    time_out = 30 #default
    dict_state_incorrect = {}
    dict_closure_incorrect = {}
    not_found_case = []
    owner_golf = {'nathawit_golf':{},'arissara_golf':{},'wisit_golf':{},'yanee_golf':{}}

    #init chrome driver
    driver_path = chromedriver_autoinstaller.install()



def login(user,pwd):

    # config
    my_config_parser = configparser.configparser()
    my_config_parser.read('config.ini')
    payload = {
        'username': my_config_parser.get('golf_login',f'{user}'),
        'password': my_config_parser.get('golf_login',f'{pwd}'),

        }


    username = payload['username']
    password = payload['password']




    global driver
    driver=webdriver.chrome(driver_path)
    driver.maximize_window()
    driver.get('https://golf.fabrinet.co.th/normaluser/myworkflow.asp?mode=1&documentgroup=&documentprefix=&docstatus=1')
    url_before = driver.current_url
    webdriverwait(driver, 10).until(ec.visibility_of_element_located((by.xpath, '//input[@type="text"]'))).send_keys(username)

    webdriverwait(driver, 10).until(ec.visibility_of_element_located((by.xpath, '//input[@type="password"]'))).send_keys(password)

    webdriverwait(driver, 10).until(ec.visibility_of_element_located((by.xpath, '//input[@type="submit"]'))).click()

    time = 0
    while(url_before==driver.current_url):
        sleep(1)
        time+=1
        print("waiting for response from server",time_out,"seconds")
        if(time > time_out):
            raise Exception("fail connected to server")




def read_excel_data():
    df = pd.read_excel('golf_8d_template.xlsm')
    print(df)
    index = df.index
    number_of_rows = len(index)
    print(number_of_rows)


    for i in range(number_of_rows):

        if 'arissaran' in df['Owner'][i].lower():
            
            owner_golf['arissara_golf'][df['golf '][i]] = [df['fa report acceptance'][i],df['cisco comment to supplier'][i],df['pid'][i]]

        elif 'natthawitp' in df['Owner'][i].lower() :

            owner_golf['nathawit_golf'][df['golf '][i]] = [df['fa report acceptance'][i],df['cisco comment to supplier'][i],df['pid'][i]]
        
        elif 'wisitl' in df['Owner'][i].lower():
            
            owner_golf['wisit_golf'][df['golf '][i]] = [df['fa report acceptance'][i],df['cisco comment to supplier'][i],df['pid'][i]]

        elif 'yaneew' in df['Owner'][i].lower() :

            owner_golf['yanee_golf'][df['golf '][i]] = [df['fa report acceptance'][i],df['cisco comment to supplier'][i],df['pid'][i]]

    print(owner_golf)
    print("="*100)


def fill_in_form(cisco_id,owner):

    frame = webdriverwait(driver, 10).until(ec.visibility_of_element_located((by.xpath, '//frame[@name="down"]')))

    driver.switch_to.frame(frame)

    webdriverwait(driver, 10).until(ec.presence_of_element_located((by.xpath, '//textarea[@name="i0s2c211"]'))).send_keys(owner_golf[owner][cisco_id][1] + ' ( pid: ' + owner_golf[owner][cisco_id][2]+' )')

    if ('final' in owner_golf[owner][cisco_id][0].lower() and 'replacement' in owner_golf[owner][cisco_id][1].lower() ) or ('final' in owner_golf[owner][cisco_id][0].lower() and 'credit' in owner_golf[owner][cisco_id][1].lower()):

        webdriverwait(driver, 10).until(ec.presence_of_element_located((by.xpath, '//input[@type="radio"][@name="i0s3c50t14"][@value="129"]'))).click()

    elif 'recall' in owner_golf[owner][cisco_id][1].lower() or 're-call' in owner_golf[owner][cisco_id][1].lower():

        webdriverwait(driver, 10).until(ec.presence_of_element_located((by.xpath, '//input[@type="radio"][@name="i0s3c50t14"][@value="130"]'))).click()
    
    elif 'final' in owner_golf[owner][cisco_id][0].lower() and 'scrap' in owner_golf[owner][cisco_id][1].lower():

        webdriverwait(driver, 10).until(ec.presence_of_element_located((by.xpath, '//input[@type="radio"][@name="i0s3c50t14"][@value="114"]'))).click()

    elif ('reject' in owner_golf[owner][cisco_id][0].lower() or 'prelim' in owner_golf[owner][cisco_id][0].lower() or 'final' in owner_golf[owner][cisco_id][0].lower()) and ('replacement' not in owner_golf[owner][cisco_id][1].lower() and 'scrap' not in owner_golf[owner][cisco_id][1].lower() and 'recall' not in owner_golf[owner][cisco_id][1].lower() and 're-call' not in owner_golf[owner][cisco_id][1].lower()  ):

        webdriverwait(driver, 10).until(ec.presence_of_element_located((by.xpath, '//input[@type="radio"][@name="i0s3c50t14"][@value="261"]'))).click()

    else:

        dict_closure_incorrect[cisco_id] = owner_golf[owner][cisco_id]

    webdriverwait(driver, 10).until(ec.visibility_of_element_located((by.xpath, '//input[@type="button"][@name="hello"][@value="send  your  action"]'))).click()

    # handle with alert
    try:
        webdriverwait(driver, 3).until(ec.alert_is_present())

        alert = driver.switch_to.alert
        alert.accept()
        print("alert accepted")

    except:
        print("no alert")

#access to the golf id


def access_to_golf_id(cisco_id,owner):
    driver.get('https://golf.fabrinet.co.th/normaluser/myworkflow.asp?mode=1&documentgroup=&documentprefix=&docstatus=1')
    webdriverwait(driver, 10).until(ec.visibility_of_element_located((by.xpath, '//input[@name="searchid"]'))).send_keys(cisco_id)
    webdriverwait(driver, 10).until(ec.visibility_of_all_elements_located((by.xpath, '//input[@name="searchbutton"]')))[0].click()

    try:
        all_td = webdriverwait(driver, 10).until(ec.visibility_of_all_elements_located((by.xpath, '//tr[@height="25"]//td')))

        print('='*200)
        all_td_text = []
        for i in all_td:
            print(i.text)
            all_td_text.append(i.text)

        print('='*200)

        print("number of spans are",len(all_td),"expected to \"11\" ")

        if all_td_text[1].strip() == cisco_id and 'fa: review' in all_td_text[8]:
            print('='*200)

            print(all_td_text[1],"is \"review state\" ")
            all_td[1].click()
            fill_in_form(cisco_id,owner)
        else:
            print('*'*100)
            # print('<<<',all_td_text[1],'<<<',cisco_id)
            print(all_td_text[1],"is not \"review state\" ")
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
                ──▓▒▒▒▓░░░ *♥finish!*♥  ░░▒▓▒▒▒▓
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
            print('show closure incorrect')
            print('-'*100)
            print(dict_closure_incorrect)
        
        if dict_state_incorrect.__len__() != 0:
            print('='*100)
            print('show state incorrect')
            print('-'*100)
            print(dict_state_incorrect)

        if not_found_case.__len__() != 0 :
            print('='*100)
            print('show not found case')
            print('-'*100)
            print(not_found_case)

        while(input("press \'e\' to exit !\n").lower() != 'e'):

            pass

        print("exit the program")
        sleep(2)
        driver.quit()
        


main()
