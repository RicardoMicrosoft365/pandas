import pandas as pd
import numpy as np
import datetime
from os import rename, listdir
import os 


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

    # Salva o DataFrame em um novo arquivo Excel
    df.to_excel("C:\\Users\\ricardo.gomes\\Downloads\\rel cwb\\" + nome_arquivo, index=False)

    # Exclui a planilha original
    os.remove("C:\\Users\\ricardo.gomes\\Downloads\\rel cwb\\CWB.csv")

    
move_file()
tratamento_dados()
