from time import sleep
from os import listdir, rename
from selenium.webdriver import Edge, Chrome
from datetime import datetime, datetime, timedelta
from dateutil.relativedelta import relativedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta


browser = Edge('C:\\Projetos\\bi\\avante\\drivers\\msedgedriver.exe')


def move_file(report: str):

    dir = 'C:\\Users\\ricardo.gomes\\Downloads\\'
    list_files = listdir(dir)
    for files in list_files:
        if ".sswweb" in files:
            rename(dir + files, dir + 'avante\\' + report + '\\' + files)
        elif ".csv" in files:

            
            rename(dir + files, dir + 'avante\\' + report + '\\' + files)

def find_windows(url: str):
    sleep(1)
    winds = browser.window_handles
    for window in winds:
        browser.switch_to.window(window)
        if url in browser.current_url:
            break

def home_page(cod: int):
    sleep(1)
    find_windows('https://sistema.ssw.inf.br/bin/menu01')
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="3"]').send_keys(cod)
    sleep(1)

def login():


    browser.get('https://sistema.ssw.inf.br/bin/ssw0422')
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="1"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="1"]').send_keys('GRT')
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="2"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="2"]').send_keys('12345678909')
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="3"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="3"]').send_keys('avante')
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="4"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="4"]').send_keys('3072!!OO')
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="5"]').send_keys(Keys.ENTER)

    
def download_156():
    sleep(2)
    find_windows('https://sistema.ssw.inf.br/bin/ssw1440')
    sleep(30)
    try:
        browser.find_element(By.LINK_TEXT, 'Atualizar').click()
        sleep(5)
        browser.find_element(By.XPATH, '/html/body/form/div[2]/div[2]/table[1]/tbody/tr[2]/td[9]/div/a/u').click()
        sleep(20)
        return True
    except:
        browser.find_element(By.LINK_TEXT, 'Atualizar').click()
        sleep(5)
        browser.find_element(By.XPATH, '//*[@id="tblsr"]/tbody/tr[22]/td[9]/div').click()
        sleep(2)
        return False
    
import datetime

def get_dates():
    # obter data atual
    now = datetime.datetime.now()

    # subtrair 1 dia para obter a data de ontem
    yesterday = now - datetime.timedelta(days=1)

    # armazenar as datas em uma lista
    dates = [("%d%m%y"), yesterday.strftime("%d%m%y")]

    return dates


def report_441():
        

        dates = get_dates()
        data_ontem = dates[1]
        print("Data de ontem:", data_ontem)

        sleep(2)
        find_windows('https://sistema.ssw.inf.br/bin/ssw0049')
        sleep(2)
        sleep(2)
        browser.find_element(By.XPATH, '//*[@id="rel_ana_per_pesq_ini"]').send_keys(Keys.CONTROL, "a")
        browser.find_element(By.XPATH, '//*[@id="rel_ana_per_pesq_ini"]').send_keys('010123')
        sleep(1)
        browser.find_element(By.XPATH, '//*[@id="rel_ana_per_pesq_fin"]').send_keys(Keys.CONTROL, "a")
        browser.find_element(By.XPATH, '//*[@id="rel_ana_per_pesq_fin"]').send_keys(dates[1])
        sleep(1)
        browser.find_element(By.XPATH, '//*[@id="rel_ana_arq_excel"]').send_keys(Keys.CONTROL, "a")
        browser.find_element(By.XPATH, '//*[@id="rel_ana_arq_excel"]').send_keys('S')
        sleep(1)
        browser.find_element(By.XPATH, '//*[@id="btn_env_rel_ana"]').send_keys(Keys.ENTER)
        sleep(1)
        browser.find_element(By.XPATH, '//*[@id="0"]').send_keys('7')
        sleep(40)
        download_156()
        move_file('441')
            


def report_200():

    # Obter a data de ontem
    # obter data atual
    now = datetime.datetime.now()
    ontem = now - timedelta(days=1)
    data_ontem = ontem.strftime("%d%m%y")

    # Obter a data de uma semana atrás
    uma_semana_atras = now - timedelta(weeks=1)
    data_semana = uma_semana_atras.strftime("%d%m%y")

    print(data_ontem, data_semana)

    lista = ['POA', 'FLN', 'CUA', 'CTB', 'CHO', 'BNU', 'JVE', 'CWB', 'LAR', 'CAC', 'MGF', 'LDB', 'PGO', 'SAO', 'CGR', 'CGB', 'ROO', 'SNO', 'GYN', 'RVD', 'ITR', 'BSB', 'SSA']
    cont = 0

    for item in lista:
        sleep(2)
        find_windows('https://sistema.ssw.inf.br/bin/ssw0495')
        sleep(2)
        browser.find_element(By.XPATH, '//*[@id="1"]').send_keys(Keys.CONTROL, "a")
        browser.find_element(By.XPATH, '//*[@id="1"]').send_keys(data_semana)
        sleep(2)
        browser.find_element(By.XPATH, '//*[@id="2"]').send_keys(Keys.CONTROL, "a")
        browser.find_element(By.XPATH, '//*[@id="2"]').send_keys(data_ontem)
        sleep(2)        
        browser.find_element(By.XPATH, '//*[@id="4"]').send_keys(Keys.CONTROL, "a")
        browser.find_element(By.XPATH, '//*[@id="4"]').send_keys(item)
        sleep(3)
        browser.find_element(By.XPATH, '//*[@id="11"]').send_keys(Keys.CONTROL, "a")
        browser.find_element(By.XPATH, '//*[@id="11"]').send_keys('E')
        sleep(5)
        if cont == 0:
            sleep(3)
            browser.find_element(By.XPATH, '//*[@id="0"]').send_keys('7')
            sleep(3)
        else:
            sleep(5)
            browser.find_element(By.XPATH, '//*[@id="12"]').send_keys(Keys.ENTER)
            sleep(3)
            browser.find_element(By.XPATH, '//*[@id="0"]').send_keys('7')

        sleep(5)
        cont = cont + 1
    
    
    


#Login OK
login()

#Home_page OK
home_page(156)

#915 baixando e exportando. 
home_page(441)
report_441()

home_page(200)
report_200()


browser.quit()


