import sys
from configuracao import Configuracao
from interface import Interface
import pandas as pd
from aposentados import Aposentados
import os
import tkinter as tk
from tkinter import filedialog
import shutil
from tela_mensagem import Mensagem as msg
from selenium_setup import SeleniumSetup
from pywinauto.application import Application
from modSiape import IniciaModSiape
import controle_terminal3270
from time import sleep
import keyboard as kb
from cdconvinc import CDCONVINC
from cacoaposse import CACOAPOSSE
import re
from datetime import datetime
from browser import Browser
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import Select
from seleciona_unidade_sei import SelecionaUnidade
from decipyx.web_automation import WebAutomation as wauto

class Vinculo():

    def __init__(self, orgao, data_aposentadoria, data_ingresso, data_exclusao, anterior, atual):
        self.orgao          = orgao
        self.aposentadoria  = data_aposentadoria
        self.ingresso       = data_ingresso
        self.exclusao       = data_exclusao
        self.anterior       = anterior
        self.atual          = atual

    def __str__(self):
        return (f"Órgão: {self.orgao}\n"
                f"Aposentadoria: {self.aposentadoria}\n"
                f"Ingresso: {self.ingresso}\n"
                f"Exclusão: {self.exclusao}\n"
                f"Anterior: {self.anterior}\n"
                f"Atual: {self.atual}")
    
    def is_regra_origem(self):
        # Converte as datas para objetos datetime
        aposentadoria_date = datetime.strptime(str(self.aposentadoria), '%d/%m/%Y')
        ingresso_date = datetime.strptime(str(self.ingresso), '%d/%m/%Y')
        exclusao_date = datetime.strptime(str(self.exclusao), '%d/%m/%Y')

        # Verifica se a data de aposentadoria está entre as datas de ingresso e exclusão
        return ingresso_date <= aposentadoria_date <= exclusao_date


class Decaposi():
    def __init__(self):

        from datetime import datetime
        dia_hoje = datetime.today().day
        if dia_hoje != 3:
            msg("Tratamento para:\nNão há dados para esta consulta em cacoaposse;\nUltimo órgão de origem em cacoaposse;\nCPF com duas aposentadorias em cacoaposse;\n\nverificar cpf's em chat alinhamento automação coate")
        
        self.configuracao = Configuracao() # Cria uma instancia da classe Configuracao

        self.config = self.configuracao.get_config() # recebe a variavel config que contém os parametros do arquivo json

        self.interface = Interface() # Cria uma instancia da classe Interface        

        self.window = self.interface.window # Recebe objeto UI da classe Interface        

        self.planilha = None

        self.dlg = None

        self.iniciar()

        self.evento_botoes()

        sys.exit(self.interface.app.exec_()) # inicia o loop de eventos da aplicação PyQt    

    def iniciar(self):        
        '''
        Pegar usuario e unidade digitado na interface e atualizar o self.config;
        pegar a senha'''

        user = self.interface.window.login_input.text()
        senha = self.interface.window.password_input.text()
        self.config["ultimo_acesso_user"] = user
        self.configuracao.atualiza_json(self.config)             
        
        '''self.planilha = 'base_dados_aposentados.xlsx'

        flag_existe_planilha = self.seleciona_planilha(self.planilha)        

        lista_aposentados = []
        
        if flag_existe_planilha:
            pass
        else:
            # lê a base de dados
            lista_aposentados = self.ler_base_dados()  # Beatriz
        
            # Atualiza a planilha
            self.atualiza_planilha(lista_aposentados)

        self.consultar_vinculo_decipex()  # Tem pronto

        msg("pausa")

        self.consultar_cacoaposse()  # Beatriz

        self.sair_siape()'''
        
        from credenciais import Credenciais

        cred = Credenciais()

        url = self.config["url_sei"]
        usuario = cred.user
        senha = cred.senha
        unidade = cred.unidade

        browser_sei = self.login_sei(url, usuario, senha, unidade)

        # fecha tela de aviso
        path = '/html/body/div[7]/div[2]/div[1]/div[3]/img'
        try:
            browser_sei.implicitly_wait(10)
            if browser_sei.find_element(By.XPATH, path):  
                browser_sei.find_element(By.XPATH, path).click()
        except Exception as e:
            print(f"Erro ao fechar janela de aviso: {e}")

        # Seleciona a unidade sei
        SelecionaUnidade(browser_sei, self.config["unidade_sei"])

        msg("pausa")      

        self.preencher_declaracao(browser_sei)  # Beatriz/André   
    
    
    def ler_base_dados(self):
        """
        Lê a planilha Excel com os dados iniciais: Nome, CPF, SIAPE

        Parameters:
            nome (str, optional): Nome a ser filtrado.
            cpf (str, optional): CPF a ser filtrado.
            siape (str, optional): SIAPE a ser filtrado.

        Returns:
            list: Lista de objetos Aposentados.

        Description:
            Este método realiza as seguintes ações:
            1. Procura pela planilha base_dados_solicitacao.xlsx na raiz do programa, se não encontrar solicita que o usuário selecione no computador.
            2. Lê os dados da planilha e instancia a classe Aposentados
            3. Cria uma lista com objetos da classe Aposentados
            4. Cria uma nova planilha na raiz do programa chamada "base_dados_solicitacao.xlsx" e popula com os dados tratados: "Nome, CPF e Matrícula".
            
        """
        lista = []

        try:
            # Ler o arquivo Excel e armazenar o resultado em um DataFrame
            df = pd.read_excel(self.planilha)           
            

            # Iterar sobre as linhas do DataFrame filtrado
            for _, row in df.iterrows():
                aposentado = Aposentados(
                    linha_planilha=row.name,
                    status=None,
                    status_cacoaposse=None,
                    nome=row['Nome'],
                    cpf=row['CPF'],
                    vinculo_decipex=None,  # Valor padrão, será atualizado posteriormente
                    siape=row['SIAPE'],
                    orgao_origem=None,  # Valor padrão, será atualizado posteriormente
                    data_aposentadoria="",  # Valor padrão, será atualizado posteriormente
                    fundamento_legal=None,  # Valor padrão, será atualizado posteriormente
                    dl_aposentadoria=None,  # Valor padrão, será atualizado posteriormente
                    data_dou=""  # Valor padrão, será atualizado posteriormente
                )
                lista.append(aposentado)
        
        except FileNotFoundError:
            print("Arquivo", self.planilha,"não encontrado. Por favor, verifique o nome do arquivo e sua localização.")

        return lista
    
    def consultar_vinculo_decipex(self):
        """
        Verifica se o CPF tem vinculo com Decipex.

        Parameters:
            self (object): Instância do objeto.
            

        Returns:
            None

        Description:
            Este método realiza as seguintes ações:
            1. Acessa o terminal 3270 - CDCONVINC.
            2. Verifica se o cpf tem vinculo com o Decipex.
            3. Atualiza os seguintes parâmetros da lista_aposentados: status e vinculo_decipex com True ou False.
        """
        #Cria uma lista atualizada a partir da planilha
        df = pd.read_excel(self.planilha)
        
        lista_aposentados = []

        for indice, linha in df.iterrows():
            status              = linha["Status"]
            nome                = linha['Nome']
            cpf                 = linha['CPF']
            siape               = linha['SIAPE']
            vinculo_decipex     = linha['Vínculo Decipex']

            print(vinculo_decipex)

            #Verifica se o vinculo_decipex está None, se sim instancia. 
            if str(vinculo_decipex) == "nan" and str(status) == "nan":
                aposentado = Aposentados(indice, status, None, nome, cpf, siape, vinculo_decipex, None, "", "", None, None)
                lista_aposentados.append(aposentado)

        url_siapenet = "https://www1.siapenet.gov.br/orgao/Login.do?method=inicio" 
        lista_vinculo = {"40802","40805","40806"}
        cd = CDCONVINC(url_siapenet, lista_vinculo)

        cd.iniciar_cdconvinc()
        
        for aposentado in lista_aposentados:

            cd.consultar_cpf(aposentado.cpf)

            lista_tuplas = cd.get_lista_dados(aposentado.cpf)

            lista_tupla_servidores_est02_est15 = []
            for lista in lista_tuplas:
                if lista[2] == 'servidor' and lista[6] == 'EST02' or lista[6] == 'EST15':
                    lista_tupla_servidores_est02_est15.append(lista)

            if len(lista_tupla_servidores_est02_est15) > 0:
                lista_dados = self.coletar_dados_vinculo(lista_tupla_servidores_est02_est15)

                self.rastrear_orgao_origem(lista_dados)

            aposentado.vinculo_decipex = cd.get_vinculo_decipex(aposentado.cpf)

            aposentado.status = cd.get_status_cpf(aposentado.cpf)

            self.atualiza_planilha(lista_aposentados)              
            
            sleep(2) 

        sys.exit()

    def rastrear_orgao_origem(self, lista_dados):
        num_vinculos = len(lista_dados)

        #pego o primeiro atual do objeto
        atual = lista_dados[0].atual
        atual = atual.split("/")
        atual = atual[0]

        msg(f"atual {atual}")

        

        #entro em orgao atual e




        #laço que fará as comparações
        for obj in lista_dados:
            pass

            
    
    def coletar_dados_vinculo(self, lista_tuplas):
        msg(lista_tuplas)
        lista_vinculos = [] 

        self.__app = Application().connect(title_re="^Terminal 3270.*")
        self.__dlg = self.__app.window(title_re="^Terminal 3270.*")
        self.__acesso_terminal = controle_terminal3270.Janela3270()

        sleep(0.5)
        for tupla in lista_tuplas:
            cpf_texto = tupla[0]
            self.__dlg.type_keys(cpf_texto) #cpf
            sleep(0.5)
            self.__dlg.type_keys("{ENTER}")
            sleep(0.5)
            self.__dlg.type_keys("x")
            sleep(0.5)
            self.__dlg.type_keys("{ENTER}")

            #navega até a linha
            xtab = int(tupla[4]) - 10
            if xtab > 0:
                for i in range(0, xtab):
                    self.__dlg.type_keys("{TAB}")
            
            sleep(0.5)
            self.__dlg.type_keys("x")
            sleep(0.5)
            self.__dlg.type_keys("{ENTER}")

            flag_avanca_pagina = True

            vinculo = Vinculo(tupla[5], None, None, None, None, None)

            while flag_avanca_pagina:
                #pegar da linha 7 = 0 até a 23 = 16
                for i in range(7, 24):
                    texto_aposentadoria     = r'[_\s]*(APOSENTADORIA)[_\s]*'
                    texto_ingresso          = r'[_\s]*(INGRESSO NO ORGAO)[_\s]*'
                    texto_exclusao          = r'[_\s]*(EXCLUSAO)[_\s]*'
                    texto_redistribuicao    = r'[_\s]*(REDISTRIBUICAO / REFORMA)[_\s]*'
                    texto_ultima_pagina     = r'[_\s]*(ULTIMA TELA DO VINCULO)[_\s]*'

                    texto_linha = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), i, 1, i, 80).strip()
                    
                    tem_aposentadoria = re.findall(texto_aposentadoria, texto_linha)
                    if tem_aposentadoria:
                        #msg("tem_aposentadoria")
                        data = None
                        for j in range(i, i+3):                                                    
                            texto = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), j, 1, j, 80).strip()
                            data = self.converter_data(texto)
                            vinculo.aposentadoria = data
                            if data is not None:
                                #msg(f"data da aposentadoria{data}")
                                break

                    tem_ingresso = re.findall(texto_ingresso, texto_linha)
                    if tem_ingresso:
                        data = None
                        for j in range(i, i+3):                                                    
                            texto = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), j, 1, j, 80).strip()
                            data = self.converter_data(texto)
                            vinculo.ingresso = data
                            if data is not None:
                                break
                    
                    tem_exclusao = re.findall(texto_exclusao, texto_linha)
                    if tem_exclusao:
                        data = None
                        for j in range(i, i+3):                                                    
                            texto = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), j, 1, j, 80).strip()
                            data = self.converter_data(texto)
                            vinculo.exclusao = data
                            if data is not None:
                                break

                    
                    tem_redistribuicao = re.findall(texto_redistribuicao, texto_linha)
                    if tem_redistribuicao:
                        anterior    = texto = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), (i+1), 27, (i+1), 41).strip()
                        vinculo.anterior = anterior
                        atual       = texto = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), (i+2), 27, (i+2), 41).strip()    
                        vinculo.atual = atual
                        

                    tem_ultima_pagina = re.findall(texto_ultima_pagina, texto_linha)
                    if tem_ultima_pagina:
                        self.__dlg.type_keys('{F12}')
                        sleep(1)
                        self.__dlg.type_keys('{F12}')
                        sleep(1)
                        self.__dlg.type_keys('{TAB}')
                        flag_avanca_pagina = False
                        break
                    
                if flag_avanca_pagina: self.__dlg.type_keys('{F8}')

            lista_vinculos.append(vinculo)
            print(vinculo)
            print("----------------------------------")
            print(f"regra {vinculo.is_regra_origem()}") 

                        
        return lista_vinculos

        
        '''
        sleep(0.5)
        cpf_texto = tupla[0]
        msg(cpf_texto)
        self.__dlg.type_keys(cpf_texto) #cpf
        sleep(0.5)
        self.__dlg.type_keys("{ENTER}")
        sleep(0.5)
        self.__dlg.type_keys("x")
        sleep(0.5)
        self.__dlg.type_keys("{ENTER}")

        msg('péra')
        
        for tupla in lista_tuplas: 
            
            #navega até a linha
            xtab = int(tupla[4]) - 10
            if xtab > 0:
                for i in range(0, xtab):
                    self.__dlg.type_keys("{TAB}")
            
            sleep(0.5)
            self.__dlg.type_keys("x")


            msg("essa pausa")


        
        pass
'''
    
    def consultar_cacoaposse(self):
        """
        Verifica demais dados no cacoaposse.

        Parameters:
            self (object): Instância do objeto.
            

        Returns:
            None

        Description:
            Este método realiza as seguintes ações:
            1. Acessa o terminal 3270 - CACOAPOSSE
            2. Coleta dados do CPF: orgao_origem, data_aposentadoria, fundamento_legal, Dl Aposentadoria e data_dou
            3. Atualiza os seguintes parâmetros da lista_aposentados: data_aposentadoria, fundamento_legal, Dl Aposentadoria e data_dou
            4. Fecha o terminal 3270
            5. atualiza a planilha excel acrescentando as devidas colunas
        """

        #Cria uma lista atualizada a partir da planilha
        df = pd.read_excel(self.planilha)
        
        lista_aposentados = []

        colunas_verificar = ['Status Cacoaposse', 'Dl Aposentadoria', 'Data Aposentadoria', 'Fundamento Legal', 'Data Publicação DOU']
        for coluna in colunas_verificar:
            if coluna not in df.columns:
                df[coluna] = ""

        for indice, linha in df.iterrows():
            status              = linha['Status']
            status_cacoaposse   = linha["Status Cacoaposse"]
            cpf                 = linha['CPF']
            vinculo_decipex     = linha['Vínculo Decipex']
            siape               = linha['SIAPE']
            dl_aposentadoria    = linha['Dl Aposentadoria']
            data_dou            = linha['Data Publicação DOU']
            data_aposentadoria  = linha['Data Aposentadoria']
            fundamento_legal    = linha['Fundamento Legal']
            nome                = linha['Nome']


            #Verifica se o status ao consultar o cacoaposse está None, se sim instancia.
            if pd.isna(status_cacoaposse):
                aposentado = Aposentados(
                    indice,
                    status            = status,
                    status_cacoaposse = status_cacoaposse,
                    nome              = nome,
                    cpf               = cpf,
                    siape             = siape,
                    vinculo_decipex   = vinculo_decipex,
                    orgao_origem      = None,
                    data_aposentadoria= str(data_aposentadoria),  
                    data_dou          = str(data_dou),  
                    fundamento_legal  = str(fundamento_legal),  
                    dl_aposentadoria  = str(dl_aposentadoria))  
                
                #condição que determina que a lista só será preenchida se status_cacoaposse for None
                if status_cacoaposse is not None:
                    lista_aposentados.append(aposentado)

        url_siapenet = "https://www1.siapenet.gov.br/orgao/Login.do?method=inicio" 
        cacoaposse = CACOAPOSSE(url_siapenet)

        cacoaposse.iniciar_cacoaposse()
             
        for aposentado in lista_aposentados:
            
            try:
                cacoaposse.consultar_cpf(aposentado.cpf)

                aposentado.status_cacoaposse    = cacoaposse.get_status_cacoaposse()

                aposentado.dl_aposentadoria     = cacoaposse.get_dl_aposentadoria()

                aposentado.set_data_dou(cacoaposse.get_data_dou())

                aposentado.set_data_aposentadoria(cacoaposse.get_data_aposentadoria())

                aposentado.fundamento_legal     = cacoaposse.get_fundamento_legal()

                self.atualiza_planilha(lista_aposentados)
            
            except Exception as erro:
                print("------------------------------->ERRO ao consultar o cpf",aposentado.cpf)
                continue
            sleep(2)
                    
    def preencher_declaracao(self, browser_sei):
        """
        Preenche declaração no SEI.

        Parameters:
            self (object): Instância do objeto.
            

        Returns:
            None

        Description:
            Este método realiza as seguintes ações:
            1. Acessa o SEI
            2. Preenche declaração
            3. Atualiza os seguintes parâmetros da lista_aposentados: data_aposentadoria, fundamento_legal, num_portaria e data_dou
            4. Fecha o terminal 3270
            5. Renomeia a planilha para "Solicitação_finalizada.xlsx
        """
        
        



    def login_sei(self, url, usuario, senha, unidade):
        browser = Browser().get_browser_01()
        try:
            browser.get(url)
            browser.maximize_window()
            sleep(3)
            browser.find_element(by=By.ID, value='selOrgao').send_keys(unidade) ##aqui
            browser.find_element(by=By.ID, value='txtUsuario').send_keys(usuario)
            browser.find_element(by=By.ID, value='pwdSenha').send_keys(senha)
            try: # nem sempre precisa clicar no botao "Acessar" para que o login aconteca
                browser.find_element(by=By.ID, value='Acessar').click()
                sleep(1)
            except:
                sleep(1)
        except self.exception as e:
            print(f"Erro ao acessar o site {url}: {e}")

        return browser

    def evento_botoes(self):
        self.window.button_iniciar.clicked.connect(self.iniciar)

    def seleciona_planilha(self, planilha):
        flag_existe_planilha = False

        if os.path.exists(planilha):
           flag_existe_planilha = True

        else:
            root = tk.Tk()
            root.withdraw()  # Esconde a janela principal do Tkinter
            file_path = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Arquivos Excel", "*.xlsx")])
            root.destroy()
            if file_path:
                            
                new_file_path = os.path.join(os.getcwd(), planilha)
                shutil.copy(file_path, new_file_path)
                print(f"Arquivo salvo como {new_file_path}")
                self.planilha = os.path.basename(new_file_path)
                
        return flag_existe_planilha

    def sair_siape(self) :
        try:
            self.app = Application().connect(title_re="^Painel de controle.*")
            self.dlg = self.app.window(title_re="^Painel de controle.*")
            self.dlg.type_keys('%{F4}')
            sleep(1)
            kb.press("Enter")
        except Exception as e:
            print(e)  
            
    def converter_data(self, texto):
        padrao_data = r'DATA OCORRENCIA:(\d{2})(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)(\d{4})'
        match = re.search(padrao_data, texto)
        
        if not match:
            return None

        # Se o padrão for encontrado, extraia as partes da data
        dia = match.group(1)
        mes_abrev = match.group(2)
        ano = match.group(3)

        # Dicionário para mapear os meses abreviados para números
        meses = {
            "JAN": "01",
            "FEV": "02",
            "MAR": "03",
            "ABR": "04",
            "MAI": "05",
            "JUN": "06",
            "JUL": "07",
            "AGO": "08",
            "SET": "09",
            "OUT": "10",
            "NOV": "11",
            "DEZ": "12"
        }
        
        # Converte o mês abreviado para número
        mes = meses.get(mes_abrev.upper(), "00")
        
        # Formata a data no formato desejado
        data_formatada = f"{dia}/{mes}/{ano}"
        
        return data_formatada
    
    def atualiza_planilha(self, lista_aposentados):

        # Abre a planilha Excel uma vez para atualização, especificando que todas as colunas devem ser lidas como strings
        base_dados_atualizada = pd.read_excel(self.planilha, dtype=str)

        # Verifica se as colunas existem na planilha, se não, as cria.
        colunas_necessarias = ['Status', 'Status Cacoaposse', 'Nome', 'CPF', 'SIAPE', 'Vínculo Decipex', 'Órgão de Origem', 'Data Aposentadoria', 'Fundamento Legal', 'Portaria Número', 'Data Publicação DOU']
        
        for coluna in colunas_necessarias:
            if coluna not in base_dados_atualizada.columns:
                base_dados_atualizada[coluna] = None

        # Reorganiza as colunas
        colunas_organizadas = colunas_necessarias + [coluna for coluna in base_dados_atualizada.columns if coluna not in colunas_necessarias]
        base_dados_atualizada = base_dados_atualizada[colunas_organizadas]

        # Atualiza os dados das respectivas colunas para cada aposentado na lista
        for aposentado in lista_aposentados:
            linha = aposentado.linha_planilha
            base_dados_atualizada.at[linha, 'Status'] = str(aposentado.status)
            base_dados_atualizada.at[linha, 'Status Cacoaposse'] = str(aposentado.status_cacoaposse)
            base_dados_atualizada.at[linha, 'Nome'] = str(aposentado.nome)
            base_dados_atualizada.at[linha, 'CPF'] = str(aposentado.cpf)
            base_dados_atualizada.at[linha, 'Vínculo Decipex'] = str(aposentado.vinculo_decipex)
            base_dados_atualizada.at[linha, 'SIAPE'] = str(aposentado.siape)
            base_dados_atualizada.at[linha, 'Órgão de Origem'] = str(aposentado.orgao_origem)
            base_dados_atualizada.at[linha, 'Data Aposentadoria'] = str(aposentado.data_aposentadoria)
            base_dados_atualizada.at[linha, 'Fundamento Legal'] = str(aposentado.fundamento_legal)
            base_dados_atualizada.at[linha, 'Portaria Número'] = str(aposentado.dl_aposentadoria)
            base_dados_atualizada.at[linha, 'Data Publicação DOU'] = str(aposentado.data_dou)

        # Salva a planilha Excel atualizada, convertendo todos os dados para string
        base_dados_atualizada = base_dados_atualizada.astype(str)
        base_dados_atualizada.to_excel(self.planilha, index=False)

    

if __name__ == "__main__":
        bot = Decaposi()
        sys.exit()