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
import schedule


class MonitorPreco:

    def __init__(self):
        self.produto = ''
        self.lista_nome_produto = []
        self.lista_lojas = []
        self.lista_precos = []
        self.lista_links = []

    def tela_pesquisa(self):
        '''
            Layout Caixa de Pesquisa
        '''
        sg.theme('Reddit')

        layout = [
        [sg.Text('Pesquisar: ', size=(12, 1)), sg.Input(key='pesquisar', size=(15, 1))],
        [sg.Button(button_text='Enter')],
        [sg.Text(key='pesq')]
        ]

        window = sg.Window('Pesquisa Google', layout=layout)

        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED:
                break

            elif event == 'Enter':
                self.produto = values['pesquisar']
                self.iniciar_driver()
                self.acessar_pagina_pesquisar()
                self.planilhar()
                break

    def iniciar_driver(self):
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

        self.driver = webdriver.Chrome(options=chrome_options)

        self.wait = WebDriverWait(
            self.driver,
            10,
            poll_frequency=1,
            ignored_exceptions=[
                NoSuchElementException,
                ElementNotVisibleException,
                ElementNotSelectableException,
            ]
        )

    def acessar_pagina_pesquisar(self):
        '''
        Acessa a página e realiza a pesquisa.
        '''
        try:
            self.driver.get('https://www.google.com.br/')
            self.driver.maximize_window()
            sleep(3)
        except:
            print('Não foi possível abrir o Navegador')

        try:    
            campo_pesquisar = self.driver.find_element(By.ID, 'APjFqb')
            campo_pesquisar.send_keys(self.produto)
            campo_pesquisar.send_keys(Keys.ENTER)
            sleep(5)
        except:
            print('Não foi possível digitar o produto na aba do Google')
            
        try:
            print('Clicando em Shopping')
            xpath_shopping = '//div[@role="listitem"]'
            abas_google = self.wait.until(expected_conditions.visibility_of_any_elements_located((
                By.XPATH, xpath_shopping)))

            for aba in abas_google:
                if aba.text == 'Shopping':
                    aba.click()
        except:
            pass

        '''
        Localiza e salvar as informações dos produtos.
        '''
        sleep(5)

        try:
            print('Salvando os Nomes dos Produtos')
            xpath_nome_produto = self.driver.find_elements(By.XPATH, "//h3[@class='tAxDx']")
            for nome_produto in xpath_nome_produto:
                self.lista_nome_produto.append(nome_produto.text)
        except:
            print('Não foi possível salvar os nomes dos produtos')
            pass

        try:
            print('Salvando os Preços')
            xpath_preco = self.driver.find_elements(By.CLASS_NAME, 'a8Pemb.OFFNJ')
            for preco in xpath_preco:
                preco = preco.text.split(' ')[1]
                preco = preco.replace('.', '')
                preco = preco.replace(',', '.')
                self.lista_precos.append(float(preco))
        except:
            print('Não foi possível salvar os preços dos produtos')
            pass

        try:
            print('Salvando as Lojas')
            xpath_loja = self.driver.find_elements(By.CLASS_NAME, 'aULzUe.IuHnof')
            for loja in xpath_loja:
                self.lista_lojas.append(loja.text)
        except:
            print('Não foi possível salvar as lojas dos produtos')
            pass

        try:
            print('Salvando os Links')
            xpath_link = self.driver.find_elements(By.XPATH, '//a[@class="shntl"]')
            condição = True
            for link in xpath_link:
                if condição == True:
                    self.lista_links.append(link.get_attribute('href'))
                    condição = False
                else:
                    condição = True
        except:
            print('Não foi possível salvar os links dos produtos')
            pass

        print('Itens Salvos')
        self.driver.close()

    def planilhar(self):
        '''
        Criando a Planilha
        '''
        horario_atual = datetime.now().strftime('%d/%m/%Y %H:%M')

        try:
            planilha = openpyxl.Workbook()
            sheet = planilha.active
            sheet.append(['Produto', 'Loja', 'Valor', 'Link', 'Data da Consulta dos Preços'])
        except:
            print('Não foi possível criar o aqruivo Planilha de Preços.xlsx')
            pass

        try: 
            for i, nome_produto in enumerate(self.lista_nome_produto):
                sheet.append([nome_produto, self.lista_lojas[i], self.lista_precos[i], self.lista_links[i], horario_atual])
        except:
            print('Erro ao inserir as informações na planilha')
            pass

        planilha.save('Planilha de Preços.xlsx')
        print('Planilha Atualizada')


self = MonitorPreco()

schedule.every(30).second.do(self.tela_pesquisa)
print(f'Próximo agendamento irá ocorrer às {schedule.next_run()}')
while True:
    schedule.run_pending()
    sleep(1)
