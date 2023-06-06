from time import sleep
from os import listdir, rename
from selenium.webdriver import Edge
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import numpy as np
import datetime
from os import rename, listdir
import os 
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

browser = Edge()


def find_windows(url: str):
    sleep(1)
    winds = browser.window_handles
    for window in winds:
        browser.switch_to.window(window)
        if url in browser.current_url:
            break

def period(month: int):
    list_period = []
    now = datetime.now()
    yesterday = now - timedelta(days=1)
    list_period.append(now.strftime('%d%m%y') + " " + now.strftime('%d%m%y'))
    for x in range(month):
        diff = yesterday - timedelta(days=x*30)
        start = diff - timedelta(days=30)
        list_period.append(start.strftime('%d%m%y') + " " + diff.strftime('%d%m%y'))
    return list_period

def fill_form(datas: list):
    sleep(1)
    for x in datas:
            browser.find_element(By.XPATH, x[0]).send_keys(Keys.CONTROL, "a")
            browser.find_element(By.XPATH, x[0]).send_keys(x[1])
            sleep(1)

def home_page(cod: int):
    sleep(1)
    find_windows('https://sistema.ssw.inf.br/bin/menu01')
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="3"]').send_keys(cod)
    sleep(1)

def login():
    sleep(1)
    browser.get('https://sistema.ssw.inf.br/bin/ssw0422')
    login = [['//*[@id="1"]','GRT'], ['//*[@id="2"]','11139703919'], ['//*[@id="3"]','ricardo'], ['//*[@id="4"]','160199']]
    fill_form(login)
    browser.find_element(By.XPATH, '//*[@id="5"]').send_keys(Keys.ENTER)

def download_156():

    sleep(2)
    find_windows('https://sistema.ssw.inf.br/bin/ssw1440')
    sleep(30)
    try:
        browser.find_element(By.LINK_TEXT, 'Atualizar').click()
        sleep(5)
        browser.find_element(By.XPATH, '/html/body/form/div[2]/div[2]/table[1]/tbody/tr[2]/td[9]/div/a/u').click()
        sleep(10)
        return True
    except:
        browser.find_element(By.LINK_TEXT, 'Atualizar').click()
        sleep(5)
        browser.find_element(By.XPATH, '//*[@id="tblsr"]/tbody/tr[22]/td[9]/div').click()
        sleep(2)
        return False

def get_dates():
    # obter data atual
    now = datetime.datetime.now()

    # subtrair 1 dia para obter a data de ontem
    yesterday = now - datetime.timedelta(days=1)

    # armazenar as datas em uma lista
    dates = [("%d%m%y"), yesterday.strftime("%d%m%y")]

    return dates

def report_455():

    dates = get_dates()
    data_ontem = dates[1]
    print("Data de ontem:", data_ontem)

    find_windows('https://sistema.ssw.inf.br/bin/ssw0230')


    browser.find_element(By.XPATH, '//*[@id="2"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="2"]').send_keys('CWB')
    sleep(2)
    browser.find_element(By.XPATH, '//*[@id="3"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="3"]').send_keys('E')    
    sleep(2)
    browser.find_element(By.XPATH, '//*[@id="9"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="9"]').send_keys(dates[1])
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="10"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="10"]').send_keys(dates[1])
    sleep(2)
    browser.find_element(By.XPATH, '//*[@id="11"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="11"]').send_keys(Keys.DELETE)
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="12"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="12"]').send_keys(Keys.DELETE)
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="35"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="35"]').send_keys('E')
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="37"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="37"]').send_keys('A')
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="38"]').send_keys(Keys.CONTROL, "a")
    browser.find_element(By.XPATH, '//*[@id="38"]').send_keys('B')
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="40"]').send_keys(Keys.ENTER)
    sleep(1)
    browser.find_element(By.XPATH, '//*[@id="0"]').send_keys('7')
    sleep(180)
    download_156()
    
def move_file():

    dir = 'C:\\Users\\ricardo.gomes\\Downloads\\'
    list_files = listdir(dir)
    for files in list_files:
        if ".sswweb" in files:
            rename(dir + files, dir + 'rel cwb\\' + 'CWB' + '.csv')
        elif ".csv" in files:

            
            rename(dir + files, dir + 'rel cwb\\' + 'CWB' + '.csv')

def tratamento_dados():

    # Lê o arquivo CSV original
    df = pd.read_csv('C:\\Users\\ricardo.gomes\\Downloads\\rel cwb\\CWB.csv', sep=';', decimal=',', encoding='ISO-8859-1', skiprows=1)

    # Lista com os números das colunas que devem ser excluídas
    cols_to_drop = [0,4,5,8,9,10,11,14,17,18,19,20,21,22,23,24,25,26,27,29,30,31,32,34,35,37,38,39,40,41]

    # Exclusão das colunas

    # Exclui as colunas do DataFrame
    df = df.drop(df.columns[cols_to_drop], axis=1)

    # Lista com as cidades a serem excluídas
    cidades_excluir = ['ADRIANOPOLIS', 'AGUDOS DO SUL', 'ALMIRANTE TAMANDARE', 'ANTONINA', 'ANTONIO OLINTO', 'ARAUCARIA', 'BALSA NOVA', 'BELA VISTA DO TOLDO', 'BITURUNA', 'BOCAIUVA DO SUL', 'CAMPINA GRANDE DO SUL', 'CAMPO ALEGRE', 'CAMPO DO TENENTE', 'CAMPO LARGO', 'CAMPO MAGRO', 'CANOINHAS', 'CERRO AZUL', 'COLOMBO', 'CONTENDA', 'CRUZ MACHADO', 'CURITIBA', 'DOUTOR ULYSSES', 'FAZENDA RIO GRANDE', 'GENERAL CARNEIRO', 'GUARAQUECABA', 'GUARATUBA', 'IRINEOPOLIS', 'ITAIOPOLIS', 'ITAPERUCU', 'LAPA', 'MAFRA', 'MAJOR VIEIRA', 'MANDIRITUBA', 'MATINHOS', 'MATOS COSTA', 'MONTE CASTELO', 'MORRETES', 'PAPANDUVA', 'PARANAGUA', 'PAULA FREITAS', 'PAULO FRONTIN', 'PIEN', 'PINHAIS', 'PIRAQUARA', 'PONTAL DO PARANA', 'PORTO AMAZONAS', 'PORTO UNIAO', 'PORTO VITORIA', 'QUATRO BARRAS', 'QUITANDINHA', 'RIO BRANCO DO SUL', 'RIO NEGRINHO', 'RIO NEGRO', 'SAO BENTO DO SUL', 'SAO JOAO DO TRIUNFO', 'SAO JOSE DOS PINHAIS', 'SAO MATEUS DO SUL', 'TIJUCAS DO SUL', 'TRES BARRAS', 'TUNAS DO PARANA', 'UNIAO DA VITORIA']

    # Obtendo os índices das linhas com as cidades a serem excluídas
    linhas_excluir = df[df['Cidade de Entrega'].isin(cidades_excluir)].index

    # Excluindo as linhas do DataFrame original
    df.drop(linhas_excluir, inplace=True)

    # Selecionar as linhas em que a coluna L não é nula
    linhas_a_excluir = df[df['Data da Entrega Realizada'].notnull()]

    # Excluir as linhas selecionadas
    df.drop(linhas_a_excluir.index, inplace=True)

    # Selecionar as linhas em que a coluna I não é nula
    linhas_a_excluir = df[df['Data do Primeiro Manifesto'].notnull()]

    # Excluir as linhas selecionadas
    df.drop(linhas_a_excluir.index, inplace=True)

    # renomeia a coluna "Unnamed: 42" para "TRATATIVA"
    df = df.rename(columns={'Unnamed: 42': 'TRATATIVA'})
    
    # Obtendo a data de ontem
    ontem = datetime.date.today() - datetime.timedelta(days=1)

    # Substitua 'CWB' pelo nome da sua cidade
    nome_arquivo = f'CWB {ontem.strftime("%d-%m-%Y")}.xlsx'

    caminho_arquivo = "C:\\Users\\ricardo.gomes\\Downloads\\rel cwb\\" + nome_arquivo

    # Salva o DataFrame em um novo arquivo Excel
    df.to_excel("C:\\Users\\ricardo.gomes\\Downloads\\rel cwb\\" + nome_arquivo, index=False)


    # Exclui a planilha original
    os.remove("C:\\Users\\ricardo.gomes\\Downloads\\rel cwb\\CWB.csv")

    print(df)

def email():

    # Obtendo a data de ontem
    ontem = datetime.date.today() - datetime.timedelta(days=1)

    # Substitua 'CWB' pelo nome da sua cidade
    nome_arquivo = f'CWB {ontem.strftime("%d-%m-%Y")}.xlsx'


    # Salva o DataFrame no formato CSV brasileiro
    caminho_arquivo = ('C:\\Users\\ricardo.gomes\\Downloads\\rel cwb\\' + nome_arquivo)
    

    # Configurações do servidor SMTP
    servidor_smtp = "smtp.gritsch.com.br"
    porta_smtp = 587
    usuario_smtp = "naoresponda@gritsch.com.br"
    senha_smtp = "N@oR3spond@"

    #login servidor
    server = smtplib.SMTP(servidor_smtp, porta_smtp)

    #start TLS
    server.ehlo()
    server.starttls()

    try:
        #login email 
        server.login(usuario_smtp, senha_smtp)
        print("Login bem-sucedido!")
    except smtplib.SMTPAuthenticationError:
        print("Erro: falha na autenticação do servidor SMTP.")

    corpo_email = "Isso é um de envia TRANFERENCIA CWB DIÁRIA!"
    
    email_teste = "ricardo.gomes@gritsch.com.br"
    email_copia = "ayrton.souza@gritsch.com.br"
    email_msg = MIMEMultipart()
    email_msg['From'] = usuario_smtp    
    email_msg['To'] = email_teste
    email_msg['Cc'] = email_copia  # adiciona um cabeçalho CC
    email_msg['Subject'] = "Email Teste"
    email_msg.attach(MIMEText(corpo_email, 'html' ) )

    # Anexa o arquivo Excel ao e-mail
    attachment = open(caminho_arquivo, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % os.path.basename(caminho_arquivo))
    email_msg.attach(part)
    
    # envia msg
    destinatarios = [email_msg['To'], email_msg['Cc']]
    server.sendmail(email_msg['From'], destinatarios, email_msg.as_string())
        
    #fecha servidor 
    server.quit



#login()
#home_page(156)
#home_page(455)
#report_455()
#browser.quit()
move_file()
tratamento_dados()
#email()