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
            #lê a base de dados
            lista_aposentados = self.ler_base_dados() #Beatriz
        
            #Atualiza a planilha
            self.atualiza_planilha(lista_aposentados)

        self.consultar_vinculo_decipex() #Tem pronto

        self.consultar_cacoaposse() #Beatriz

        self.preencher_declaracao() #Beatriz/André       
    
    
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
                    nome=row['Nome'],
                    cpf=row['CPF'],
                    vinculo_decipex=None,  # Valor padrão, será atualizado posteriormente
                    siape=row['SIAPE'],
                    orgao_origem=None,  # Valor padrão, será atualizado posteriormente
                    data_aposentadoria=None,  # Valor padrão, será atualizado posteriormente
                    fundamento_legal=None,  # Valor padrão, será atualizado posteriormente
                    num_portaria=None,  # Valor padrão, será atualizado posteriormente
                    data_dou=None  # Valor padrão, será atualizado posteriormente
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
            1. Acessa o terminal 3270 - CDCONVINC
            2. Verifica se o cpf tem vinculo com o Decipex
            3. Atualiza os seguintes parâmetros da lista_aposentados: status e vinculo_decipex com True ou False
            4. atualiza a planilha excel acrescentando vinculo_decipex
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
                aposentado = Aposentados(indice, status, nome, cpf, siape, vinculo_decipex, None, None, None, None, None)
                lista_aposentados.append(aposentado)

        for aposentado in lista_aposentados:
            print(aposentado.linha_planilha)

                    

        #se o terminal 3270 (tela preta) estiver aberto, fecha.
        try :
            self.app = Application().connect(title_re="^Painel de controle.*")
            self.sair_siape()
            
        except Exception as e :
            pass   
        
        siape = SeleniumSetup()
        IniciaModSiape(siape, self.json["url_siapenet"]).executar_siape()
        
        #o bloco abaixo fica aguardando a tela preta aparecer por 30 segundos
        contador_telapreta = 0
        flag_telapreta = True

        while flag_telapreta :
            try :
                self.app = Application().connect(title_re="^Terminal 3270.*")
                self.dlg = self.app.window(title_re="^Terminal 3270.*")
                self.acesso = controle_terminal3270.Janela3270()
                flag_telapreta = False
            except Exception as e :
                contador_telapreta+=1
                print("Tentativa",contador_telapreta)
                if contador_telapreta >= 30 :
                    flag_telapreta = False
                    print("Falha ao abrir SiapeNet")
            sleep(1)
            
        sleep(2)

        
        
        #faz o for de cada linha da lista (cada aposentado) e consulta se tem vinculo
        for aposentado in lista_aposentados:
            print("-------------------------------")
            print("-------------------------------")
            print("-------------------------------")
            print("-------------------------------")
            print("-------------------------------")
            print("Consultando:",aposentado.cpf)
            print("-------------------------------")
            print("-------------------------------")
            print("-------------------------------")
            print("-------------------------------")
            print("-------------------------------")

            tela = cpf_mensagem = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 1, 1, 24, 80).strip()
                    
            #Copia a tela e verifica se há mansagens cadastradas
            partes_tela = tela.split(" ")
                        
            txt1 = False
            txt2 = False
                        
            for pt in partes_tela :
                if pt == "MENSAGEM(NS)" :
                    txt1 = True
                if pt == "CADASTRADA(S):" :
                    txt2 = True
                                
            if txt1 and txt2 :
                self.dlg.type_keys('{F3}')
                self.dlg.type_keys('{F2}')                   
                sleep(0.1)                    
                self.dlg.type_keys(">"+'CDCONVINC')
                sleep(0.1)
                kb.press("Enter")                    
                sleep(0.1)
            else :
                txt = cpf_mensagem = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 6, 6, 6, 61).strip()
                if txt == "POSICIONE O CURSOR NA OPCAO DESEJADA E PRESSIONE <ENTER>" :
                    flag_ultima_pagina = False
                    while not flag_ultima_pagina :
                        self.dlg.type_keys('{F8}')    
                        txt = cpf_mensagem = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 23, 8, 23, 20).strip()
                        if txt == "ULTIMA PAGINA" :
                            sleep(0.1)                    
                            self.dlg.type_keys(">"+"CDCONVINC")
                            flag_ultima_pagina = True
                        
                            sleep(0.1)
                            kb.press("Enter")                    
                            sleep(0.1)
                else:
                    self.dlg.type_keys('{F3}')
                    self.dlg.type_keys('{F2}')                   
                    sleep(0.1)                    
                    self.dlg.type_keys(">"+'CDCONVINC')
                    sleep(0.1)
                    kb.press("Enter")                    
                    sleep(0.1)
                        
            self.dlg.type_keys('{TAB}')
            sleep(0.1)
            
            self.dlg.type_keys(aposentado.cpf)                    
            sleep(0.1)
            
            kb.press("Enter")

            flag_acesso_negado = True

            while flag_acesso_negado :
                try:
                    cpf_mensagem = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 24, 9, 24, 44).strip()
                    texto_dif_cpf = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 24, 2, 24, 16).strip()
                    texto_cpf_invalido = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 24, 9, 24, 21).strip()
                    flag_acesso_negado = False                             
                except Exception as e :
                    print(e)

            if texto_cpf_invalido == "CPF INVALIDO":
                aposentado.status = "CPF inválido"
                continue
            
            elif texto_dif_cpf == "INFORME APENAS":
                aposentado.status = "CPF inválido"
                continue

            elif cpf_mensagem == "DIGITO VERIFICADOR INVALIDO":
                aposentado.status = "CPF inválido"
                continue
                            
            elif cpf_mensagem == "NAO EXISTEM DADOS PARA ESTA CONSULTA":
                aposentado.status = "CPF não cadastrado"
                continue

            else:
                self.dlg.type_keys('x')  # marca o nome

                sleep(0.1)

                kb.press("Enter")
                        
                if self.verifica_vinculo(self.json["vinculos_decipex"]):
                    aposentado.vinculo_decipex = True                    
                else:
                    aposentado.vinculo_decipex = False                    

            self.atualiza_planilha(lista_aposentados)
            
            sleep(2)


    def verifica_vinculo(self, lista_orgaos_decipex):
        pagina_linha_orgao = []
        
        # verifica a quantidade de páginas que o serv/pens possui de vinculos    
        sleep(2)
        self.qt_pagina_vinculo = self.acesso.pega_texto_siape(self.acesso.copia_tela(), 9, 78, 9, 80).strip()            

        self.qt_pagina_vinculo = int(self.qt_pagina_vinculo)
        sleep(0.1)

        for qtp in range(self.qt_pagina_vinculo):

            if self.qt_pagina_vinculo > 1: self.dlg.type_keys("{F8}")  # Se quantidade de páginas for maior que 1 tecla F8 para alternar e coletar outras páginas

            sleep(0.3)

        pagina_linha_orgao = self.popula_tupla()

        flag = False

        for org_decipex in lista_orgaos_decipex:
            
            flag = bool([(pagina, linha, orgao, tipo) for pagina, linha, orgao, tipo in pagina_linha_orgao if orgao == org_decipex])
            if flag : break
            
        return flag

    def popula_tupla(self):
        tipos = []
        pagina_linha_orgao = []

        for qtp in range(self.qt_pagina_vinculo):
            linha = 10
            linha_representativa = 0

            if self.qt_pagina_vinculo > 1: self.dlg.type_keys("{F8}")  # Se quantidade de páginas for maior que 1 tecla F8 para alternar e coletar outras páginas

            for a in range(13):  # trocar depois pra 13 que pesquisa até a linha 22
                pens = False

                pega_tipo = self.acesso.pega_texto_siape(self.acesso.copia_tela(), linha, 7, linha, 18).strip()  # Pega se é servidor ou pensionista
                tipos.append(pega_tipo)  # cria uma lista com os tipos de vínculos

                if pega_tipo == "SERVIDOR":
                    pens = False

                    pega_orgaoservidor = self.acesso.pega_texto_siape(self.acesso.copia_tela(), linha, 25, linha, 29).strip()  # Pega o órgão do servidor
                    # orgaos_na_tela.append(str(qtp)+"-"+pega_orgaoservidor)
                    pagina_linha_orgao.append((qtp, linha_representativa, pega_orgaoservidor, pens))

                elif pega_tipo == "PENSIONISTA":
                    pens = True

                    pega_orgaopensionista = self.acesso.pega_texto_siape(self.acesso.copia_tela(), linha, 40, linha, 44).strip()  # Pega o órgão do pensionista
                    # orgaos_na_tela.append(str(qtp)+"-"+pega_orgaopensionista)
                    pagina_linha_orgao.append((qtp, linha_representativa, pega_orgaopensionista, pens))

                linha += 1
                linha_representativa += 1
                
            
        return pagina_linha_orgao
    
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
            2. Coleta dados do CPF: orgao_origem, data_aposentadoria, fundamento_legal, num_portaria e data_dou
            3. Atualiza os seguintes parâmetros da lista_aposentados: data_aposentadoria, fundamento_legal, num_portaria e data_dou
            4. Fecha o terminal 3270
            5. atualiza a planilha excel acrescentando as devidas colunas
        """

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
        colunas_necessarias = ['Status', 'Nome', 'CPF', 'SIAPE', 'Vínculo Decipex', 'Órgão de Origem', 'Data Aposentadoria', 'Fundamento Legal', 'Portaria Número', 'Data Publicação DOU']
        
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
            base_dados_atualizada.at[linha, 'Nome'] = str(aposentado.nome)
            base_dados_atualizada.at[linha, 'CPF'] = str(aposentado.cpf)
            base_dados_atualizada.at[linha, 'Vínculo Decipex'] = str(aposentado.vinculo_decipex)
            base_dados_atualizada.at[linha, 'SIAPE'] = str(aposentado.siape)
            base_dados_atualizada.at[linha, 'Órgão de Origem'] = str(aposentado.orgao_origem)
            base_dados_atualizada.at[linha, 'Data Aposentadoria'] = str(aposentado.data_aposentadoria)
            base_dados_atualizada.at[linha, 'Fundamento Legal'] = str(aposentado.fundamento_legal)
            base_dados_atualizada.at[linha, 'Portaria Número'] = str(aposentado.num_portaria)
            base_dados_atualizada.at[linha, 'Data Publicação DOU'] = str(aposentado.data_dou)

        # Salva a planilha Excel atualizada, convertendo todos os dados para string
        base_dados_atualizada = base_dados_atualizada.astype(str)
        base_dados_atualizada.to_excel(self.planilha, index=False)

    

if __name__ == "__main__":
        bot = Decaposi()
        sys.exit()