from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import *
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions

import PySimpleGUI as sg
import openpyxl
from time import sleep
from datetime import datetime
import os
import schedule


def caixa_de_pesquisa():
    '''
    Layout Caixa de Pesquisa
    '''
    try:
        sg.theme('Reddit')

        layout = [
        [sg.Text('Pesquisar: ', size=(12, 1)), sg.Input(key='pesquisar', size=(15, 1))],
        [sg.Button(button_text='Enter')],
        ]

        window = sg.Window('Pesquisa Google', layout=layout)

        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED:
                break

            elif event == 'Enter':
                produto = values['pesquisar']
                break
        
        return produto
    except:
        print('Erro ao Abrir a Caixa de Pesquisa')

def iniciar_driver():
    chrome_options = Options()

    arguments = ['--lang=pt-BR', '--window-size=800,600',
                '--incognito']

    for argument in arguments:
        chrome_options.add_argument(argument)

    caminho_padrao_para_download = 'E:\\Storage\\Desktop'

    chrome_options.add_experimental_option("prefs", {
        'download.default_directory': caminho_padrao_para_download,
        'download.directory_upgrade': True,
        'download.prompt_for_download': False,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.automatic_downloads": 1,
    })

    driver = webdriver.Chrome(options=chrome_options)

    wait = WebDriverWait(
        driver,
        10,
        poll_frequency=1,
        ignored_exceptions=[
            NoSuchElementException,
            ElementNotVisibleException,
            ElementNotSelectableException,
        ]
    )

    return driver, wait

def acessar_pagina_pesquisar():
    '''
    Acessa a página e realiza a pesquisa.
    '''
    try:
        produto = caixa_de_pesquisa()
        driver, wait = iniciar_driver()

        driver.get('https://www.google.com.br/')
        driver.maximize_window()
        sleep(3)
    except:
        print('Não foi possível abrir o Navegador')

    try:    
        campo_pesquisar = driver.find_element(By.ID, 'APjFqb')
        campo_pesquisar.send_keys(produto)
        campo_pesquisar.send_keys(Keys.ENTER)
        sleep(5)
    except:
        print('Não foi possível digitar o produto na aba do Google')
        
    try:
        print('Clicando em Shopping')
        xpath_shopping = '//div[@role="listitem"]'
        abas_google = wait.until(expected_conditions.visibility_of_any_elements_located((
            By.XPATH, xpath_shopping)))

        for aba in abas_google:
            if aba.text == 'Shopping':
                aba.click()
    except:
        print('Erro ao Clicar na aba Shopping')

    '''
    Localiza e salvar as informações dos produtos.
    '''
    sleep(5)

    lista_nome_produto = []
    lista_lojas = []
    lista_precos = []
    lista_links = []

    try:
        print('Salvando os Nomes dos Produtos')
        xpath_nome_produto = driver.find_elements(By.XPATH, "//h3[@class='tAxDx']")
        for nome_produto in xpath_nome_produto:
            lista_nome_produto.append(nome_produto.text)
    except:
        print('Não foi possível salvar os nomes dos produtos')
        pass

    try:
        print('Salvando os Preços')
        xpath_preco = driver.find_elements(By.CLASS_NAME, 'a8Pemb.OFFNJ')
        for preco in xpath_preco:
            preco = preco.text.split(' ')[1]
            preco = preco.replace('.', '')
            preco = preco.replace(',', '.')
            lista_precos.append(float(preco))
    except:
        print('Não foi possível salvar os preços dos produtos')
        pass

    try:
        print('Salvando as Lojas')
        xpath_loja = driver.find_elements(By.CLASS_NAME, 'aULzUe.IuHnof')
        for loja in xpath_loja:
            lista_lojas.append(loja.text)
    except:
        print('Não foi possível salvar as lojas dos produtos')
        pass

    try:
        print('Salvando os Links')
        xpath_link = driver.find_elements(By.XPATH, '//a[@class="shntl"]')
        condição = True
        for link in xpath_link:
            if condição == True:
                lista_links.append(link.get_attribute('href'))
                condição = False
            else:
                condição = True
    except:
        print('Não foi possível salvar os links dos produtos')
        pass

    print('Itens Salvos')
    driver.close()

    return lista_nome_produto, lista_lojas, lista_precos, lista_links

def planilhar():
    '''
    Criando a Planilha
    '''
    lista_nome_produto, lista_lojas, lista_precos, lista_links = acessar_pagina_pesquisar()

    horario_atual = datetime.now().strftime('%d/%m/%Y %H:%M')
    try:
        caminho_da_planilha = 'Planilha de Preços.xlsx'
        if os.path.isfile(caminho_da_planilha):
            planilha = openpyxl.load_workbook(caminho_da_planilha)
            sheet = planilha.active

        else:
            planilha = openpyxl.Workbook()
            sheet = planilha.active
            sheet.append(['Produto', 'Loja', 'Valor', 'Link', 'Data da Consulta dos Preços'])
    except:
        print('Não foi possível localizar nem criar o aqruivo Planilha de Preços.xlsx')
        pass

    try: 
        for i, nome_produto in enumerate(lista_nome_produto):
            sheet.append([nome_produto, lista_lojas[i], lista_precos[i], lista_links[i], horario_atual])
    except:
        print('Erro ao inserir as informações na planilha')
        pass

    planilha.save('Planilha de Preços.xlsx')
    print('Planilha Atualizada')

def programa():
    caixa_de_pesquisa()
    iniciar_driver()
    acessar_pagina_pesquisar()
    planilhar()


schedule.every(30).minutes.do(programa)

print(f'Próximo agendamento irá ocorrer às {schedule.next_run()}')
while True:
    schedule.run_pending()
    sleep(1)