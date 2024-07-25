import sys
from configuracao import Config
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

class Decaposi():
    def __init__(self):
        self.config = Config() # Cria uma instancia da classe Config

        self.json = self.config.get_json() # Pega o config_json

        self.interface = Interface() # Cria uma instancia da classe Interface

        self.window = self.interface.window # Recebe objeto UI da classe Interface        

        self.planilha = None

        self.dlg = None

        self.evento_botoes()        

        sys.exit(self.interface.app.exec_()) # inicia o loop de eventos da aplicação PyQt    

    def iniciar(self):
        print("Iniciou")

        self.planilha = 'base_dados_aposentados.xlsx'

        flag_existe_planilha = self.seleciona_planilha(self.planilha)        

        lista_aposentados = []
        
        if flag_existe_planilha:
            pass
        else:
            # lê a base de dados
            lista_aposentados = self.ler_base_dados()  # Beatriz
        
            # Atualiza a planilha
            self.atualiza_planilha(lista_aposentados)

        #self.consultar_vinculo_decipex()  # Tem pronto

        self.consultar_cacoaposse()  # Beatriz

        self.preencher_declaracao()  # Beatriz/André    
    
    
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
                '''linha_planilha, status, status_cacoaposse, nome, cpf, siape, vinculo_decipex, orgao_origem, data_aposentadoria, data_dou, fundamento_legal, dl_aposentadoria'''
                aposentado = Aposentados(indice, status, None, nome, cpf, siape, vinculo_decipex, None, "", "", None, None)
                lista_aposentados.append(aposentado)

        url_siapenet = "https://www1.siapenet.gov.br/orgao/Login.do?method=inicio" 
        lista_vinculo = {"40802","40805","40806"}
        cd = CDCONVINC(url_siapenet, lista_vinculo)

        cd.iniciar_cdconvinc()
        
        for aposentado in lista_aposentados:

            try:

                cd.consultar_cpf(aposentado.cpf)

                aposentado.vinculo_decipex = cd.get_vinculo_decipex(aposentado.cpf)

                aposentado.status = cd.get_status_cpf(aposentado.cpf)
                
                self.atualiza_planilha(lista_aposentados)

            except Exception as erro:

                self.consultar_vinculo_decipex()
            
            sleep(2)


    
    
    
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
            status_cacoaposse   = linha["Status Cacoaposse"]
            cpf                 = linha['CPF']
            dl_aposentadoria    = linha['Dl Aposentadoria']
            data_dou            = linha['Data Publicação DOU']
            data_aposentadoria  = linha['Data Aposentadoria']
            fundamento_legal    = linha['Fundamento Legal']

            print(status_cacoaposse)

            #Verifica se o status ao consultar o cacoaposse está None, se sim instancia. 
            if pd.isna(status_cacoaposse):
                
                aposentado = Aposentados(
                    indice,
                    status=None,
                    status_cacoaposse=status_cacoaposse,
                    nome="",
                    cpf=cpf,
                    siape="",
                    vinculo_decipex=None,
                    orgao_origem=None,
                    data_aposentadoria=str(data_aposentadoria),  
                    data_dou=str(data_dou),  
                    fundamento_legal=str(fundamento_legal),  
                    dl_aposentadoria=str(dl_aposentadoria))  
                
                lista_aposentados.append(aposentado)

        url_siapenet = "https://www1.siapenet.gov.br/orgao/Login.do?method=inicio" 
        cd = CACOAPOSSE(url_siapenet)

        cd.iniciar_cacoaposse()
             
        for aposentado in lista_aposentados:

            cd.consultar_cpf(aposentado.cpf)
            aposentado.status_cacoaposse    = cd.get_status_cacoaposse()

            aposentado.consultar_cpf        = cd.consultar_cpf(aposentado.cpf)

        self.atualiza_planilha(lista_aposentados)

    def preencher_declaracao(self):
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