import sys
from configuracao import Configuracao
from interface import Interface
import pandas as pd
from aposentado import Aposentado
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
import re
from datetime import datetime
from browser import Browser
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import Select
from seleciona_unidade_sei import SelecionaUnidade
from decipyx.web_automation import WebAutomation as wauto
from tkinter import simpledialog

class Vinculo():

    def __init__(self):
        self.orgao                      = None
        self.is_elegivel_orgao_origem   = False
        self.aposentadoria              = None
        self.ingresso                   = None
        self.exclusao                   = None
        self.orgao_anterior             = None
        self.matricula_anterior         = None
        self.orgao_atual                = None
        self.matricula_atual            = None
        self.is_linkado                 = False
        self.is_visitado                = False
    
    def is_regra_origem(self):
        """
        Verifica se a data de aposentadoria está entre as datas de ingresso e exclusão.

        Esta função converte as datas de aposentadoria, ingresso e exclusão para objetos
        `datetime` e verifica se a data de aposentadoria está no intervalo entre a data de ingresso
        e a data de exclusão, inclusive.

        Retorna:
            bool: Retorna True se a data de aposentadoria estiver entre as datas de ingresso e exclusão. 
            Retorna False se qualquer uma das datas (aposentadoria, ingresso ou exclusao) for None ou 
            se a data de aposentadoria estiver fora do intervalo.

        Exceções:
            Levanta uma exceção ValueError se as datas estiverem em um formato inválido.

        """
        if self.aposentadoria is not None and self.ingresso is not None and self.exclusao is not None: 
            aposentadoria_date = datetime.strptime(str(self.aposentadoria), '%d/%m/%Y')
            ingresso_date = datetime.strptime(str(self.ingresso), '%d/%m/%Y')
            exclusao_date = datetime.strptime(str(self.exclusao), '%d/%m/%Y')

            # Verifica se a data de aposentadoria está entre as datas de ingresso e exclusão
            return ingresso_date <= aposentadoria_date <= exclusao_date
        else:
            return False
            
    def __str__(self):
        return (f"Órgão: {self.orgao}\n"
                f"is_elegivel_orgao_origem: {self.is_elegivel_orgao_origem}\n"
                f"Aposentadoria: {self.aposentadoria}\n"
                f"Ingresso: {self.ingresso}\n"
                f"Exclusão: {self.exclusao}\n"
                f'Órgao_anterior: {self.orgao_anterior}\n'
                f'Matricula_anterior: {self.matricula_anterior}\n' 
                f'Órgao_atual: {self.orgao_atual}\n'
                f'Matricula_atual: {self.matricula_atual}\n'
                f'is_visitado: {self.is_visitado}\n'                
        )


class Decaposi():
    def __init__(self):

        from datetime import datetime
        dia_hoje = datetime.today().day
        '''if dia_hoje != 27:
            msg("Tratamento para:\nCriar regra que fique olhando para tela preta e se não houver fechar o programa\nQuando for colocar o orgão de origem colocar o codigo - nome_do_orgão;\nOs que forem 17000 devem ter os dois nomes com a regra antes e depois de 2019;\nNão há dados para esta consulta em cacoaposse;\nUltimo órgão de origem em cacoaposse;\nCPF com duas aposentadorias em cacoaposse;\n\nverificar cpf's em chat alinhamento automação coate")
        '''
        self.configuracao = Configuracao() # Cria uma instancia da classe Configuracao

        self.json = self.configuracao.get_json() # recebe a variavel config que contém os parametros do arquivo json

        self.numero_processo_sei = self.json["numero_processo_sei"]

        self.interface = Interface() # Cria uma instancia da classe Interface        

        self.window = self.interface.window # Recebe objeto UI da classe Interface

        self.window.login_input.setText(self.json["ultimo_acesso_user"])       

        self.planilha = None

        self.dlg = None

        self.evento_botoes()

        sys.exit(self.interface.app.exec_()) # inicia o loop de eventos da aplicação PyQt '''

    # Função para solicitar o número do processo SEI
    def pedir_numero_processo(self):
        # Criação da janela principal (oculta)
        root = tk.Tk()
        root.withdraw()  # Esconde a janela principal

        valor_padrao = self.json["numero_processo_sei"]

        # Janela de diálogo para capturar o número do processo SEI
        numero_processo_sei = simpledialog.askstring("ATENÇÃO!", "Verifique se o número do processo SEI está correto.\nClique OK ou se preferir digitre o número correto.",
                                                 initialvalue=valor_padrao)

        # Exibe o número do processo SEI digitado no console
        if numero_processo_sei:
            print(f"O número do processo SEI digitado foi: {numero_processo_sei}")
        else:
            print("Nenhum número de processo foi digitado.")

        # Fecha e destrói a janela principal
        root.destroy()

        return numero_processo_sei

    def iniciar(self):

        self.numero_processo_sei = self.pedir_numero_processo()

        if self.numero_processo_sei is not None:
            self.json["numero_processo_sei"] = self.numero_processo_sei
            self.configuracao.atualiza_json(self.json)
        else:
            sys.exit()

        self.planilha = 'planilha_log.xlsx'
        self.tabela_orgaos = self.json["caminho_tabela_orgaos"]    
        self.ler_base_dados()

        user = self.interface.window.login_input.text()
        senha = self.interface.window.password_input.text()
        '''from credenciais import Credenciais
        cred = Credenciais()
        senha = cred.senha'''
        unidade = self.interface.window.unidade_combo_box.currentText()
 
        self.json["ultimo_acesso_user"] = user
        self.json["ultimo_orgao"] = unidade

        self.configuracao.atualiza_json(self.json)
        
        url = self.json["url_sei"]
        usuario = user
        senha = senha
        unidade = unidade
 
        self.browser_sei = self.login_sei(url, usuario, senha, unidade)

        # fecha tela de aviso
        path = '/html/body/div[7]/div[2]/div[1]/div[3]/img'
        try:
            self.browser_sei.implicitly_wait(10)
            if self.browser_sei.find_element(By.XPATH, path):  
                self.browser_sei.find_element(By.XPATH, path).click()
        except Exception as e:
            print(f"Erro ao fechar janela de aviso: {e}")
 
        # Seleciona a unidade sei
        SelecionaUnidade(self.browser_sei, self.json["unidade_sei"])        

        self.consulta_dados_cria_declaracao()

        # Obtém a data e hora atuais no formato AAAAMMDDHH
        timestamp = datetime.now().strftime("%Y%m%d%H%M")

        # Cria o novo nome do arquivo
        novo_nome = f'planilha_log_pronto_{timestamp}.xlsx'

        # Renomeia o arquivo
        os.rename('planilha_log.xlsx', novo_nome)

        self.sair_siape()       

    def ler_base_dados(self):
        # Verifica se existe a planilha tabela de órgãos, se não, pede pro usuário selecioná-la.
        if not os.path.exists(self.tabela_orgaos):
            msg("Selecione a planilha que contém a tabela de órgãos.")
            root = tk.Tk()
            root.withdraw()  # Esconde a janela principal do Tkinter
            file_path = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Arquivos Excel", "*.xlsx")])
            self.tabela_orgaos = os.path.basename(file_path)
            self.json["caminho_tabela_orgaos"] = self.tabela_orgaos
            self.configuracao.atualiza_json(self.json)
            root.destroy()

        # Verifica se existe a planilha tabela de CPFs, se não, pede pro usuário selecioná-la.        
        if not os.path.exists(self.planilha):
            msg("Selecione a planilha que contém a lista de CPF.")          
            root = tk.Tk()
            root.withdraw()  # Esconde a janela principal do Tkinter
            file_path = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Arquivos Excel", "*.xlsx")])
            root.destroy()

            if file_path:                            
                new_file_path = os.path.join(os.getcwd(), self.planilha)
                shutil.copy(file_path, new_file_path)                
                self.planilha = os.path.basename(new_file_path)            

            lista = []

            try:
                # Ler o arquivo Excel e garantir que a coluna CPF seja lida como string
                df = pd.read_excel(self.planilha, dtype={'CPF': str})

                # Limpar os CPFs removendo os pontos e traços (caso existam) e mantendo os zeros à esquerda
                df['CPF'] = df['CPF'].str.replace(".", "").str.replace("-", "")

                # Iterar sobre as linhas do DataFrame filtrado
                for _, row in df.iterrows():
                    aposentado = Aposentado(
                        linha_planilha      =row.name,
                        status              =None,                    
                        cpf                 =row['CPF'],
                        nome                =None,
                        siape               =None,
                        vinculo_decipex     =None,                   
                        orgao_origem        =None, 
                        data_aposentadoria  =None, 
                        data_dou            =None, 
                        fundamento_legal    =None, 
                        dl_aposentadoria    =None, 
                    )
                    lista.append(aposentado)
            
            except FileNotFoundError:
                print("Arquivo", self.planilha, "não encontrado. Por favor, verifique o nome do arquivo e sua localização.")
            
            self.atualiza_planilha(lista)
    
    def consulta_dados_cria_declaracao(self):
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

        df['CPF'] = df['CPF'].astype(str).str.replace(".", "").str.replace("-", "").str.zfill(11)
        
        lista_aposentados = []

        for _, row in df.iterrows():
            aposentado = Aposentado(
                linha_planilha      =row.name,                
                status              =row["Status"],                                   
                cpf                 =row['CPF'], 
                nome                =row['Nome'],              
                siape               =row['SIAPE'],
                vinculo_decipex     =row['Vínculo Decipex'],
                orgao_origem        =row['Órgão de Origem'],
                data_aposentadoria  =row['Data Aposentadoria'],
                data_dou            =row['Data Publicação DOU'],                
                fundamento_legal    =row['Fundamento Legal'],
                dl_aposentadoria    =row['Portaria Número']  
                             
            )
            
            texto = aposentado.status
            texto = str(texto).lower()
            if texto == "nan":
                lista_aposentados.append(aposentado)
            else:
                print(f"pulou {aposentado.cpf}")           

        url_siapenet    = self.json["url_siapenet"]
        lista_vinculo   = self.json["vinculos_decipex"]
        
        cd = CDCONVINC(url_siapenet, lista_vinculo)
        
        cd.iniciar_cdconvinc()

        for aposentado in lista_aposentados:
            self.browser_sei.minimize_window()
            # primeiro coleta dados no CDCONVINC
            cd.consultar_cpf(aposentado.cpf)

            if cd.get_vinculo_decipex(aposentado.cpf):
                aposentado.vinculo_decipex = "SIM"
                self.atualiza_planilha(lista_aposentados)
                pass
            else:
                aposentado.status = "Não possui vínculo com DECIPEX"
                aposentado.vinculo_decipex = "NÃO"
                self.atualiza_planilha(lista_aposentados)
                continue

            lista_tuplas = cd.get_lista_dados(aposentado.cpf)

            aposentado.nome     = lista_tuplas[0][9]
            aposentado.siape    = lista_tuplas[0][10]      

            print(f"lista_tuplas {lista_tuplas}")

            print(aposentado.cpf)
            print("linha")
            print("linha")
            agora = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
            print(agora)
            print(aposentado.cpf)

            lista_tupla_servidores_est02_est15 = []

            for lista in lista_tuplas:

                if lista[2] == 'servidor' and lista[6] == 'EST02' or lista[6] == 'EST15':
                    lista_tupla_servidores_est02_est15.append(lista)

            if len(lista_tupla_servidores_est02_est15) > 0:
                # coleta dados em cacoaposse, a saber: Fundamento Legal, Portaria Número e Data da Publicação DOU.
                status_caco = self.coletar_dados_em_cacoaposse(aposentado)

                if status_caco == "OK":
                    pass
                else:
                    aposentado.status = status_caco
                    continue

                self.atualiza_planilha(lista_aposentados)

                # coleta dados em cdconvinc, a saber: data aposentadoria, data ingresso, data exclusão e redistribuição
                lista_dados = self.coletar_dados_em_cdconvinc(lista_tupla_servidores_est02_est15)
                
                # navega sobre os objetos (vinculos) e verifica se há conexão entre eles e se atende regra
                obj = self.grafo_objetos(lista_dados)

                if obj is not None:
                    nome_orgao = self.buscar_nome_orgao(self.tabela_orgaos, str(obj.orgao))

                    #data_a_verificar recebe a data aposentadoria em formato date para testar se é anterior a 2019
                    data_a_verificar = datetime.strptime(str(obj.aposentadoria), '%d/%m/%Y')
                    if str(obj.orgao) == "17000" and data_a_verificar < 2019:
                        nome_orgao = "Ministério da Fazenda"

                    print(obj)
                    descricao_orgao = str(obj.orgao) + " - " + nome_orgao
                    aposentado.orgao_origem = descricao_orgao
                    if aposentado.data_aposentadoria != obj.aposentadoria:
                        self.log(f"Datas são diferentes para o CPF {aposentado.cpf}.\nEm cacoaposse: {aposentado.data_aposentadoria}\nEm cdconvinc: {obj.aposentadoria}")
                    aposentado.set_data_aposentadoria(obj.aposentadoria)
                    print(f"órgão de origem: {obj.orgao}")
                    
                else:
                    aposentado.orgao_origem = "Verificação manual!"
                    aposentado.status = "Não foi possível verificar órgão de origem. Verificação manual!"
                    print("Verificação Manual")            
                    
            if cd.get_vinculo_decipex(aposentado.cpf):
                aposentado.vinculo_decipex = "SIM"
            else:
                aposentado.vinculo_decipex = "NÃO"

            aposentado.status = cd.get_status_cpf(aposentado.cpf)

            self.atualiza_planilha(lista_aposentados)

            if aposentado.status == "OK":
                self.preencher_declaracao(aposentado)
                #msg("Pausa linha 400. Verificar Declaração!")

            sleep(2)

        

    def preencher_declaracao(self, aposentado):
        #self.app = Application().connect(title_re="^SEI.*")
        #self.dlg = self.app.window(title_re="^SEI.*")

        self.browser_sei.maximize_window()
 
        elemento_caixa_pesquisar = '//*[@id="txtPesquisaRapida"]'
        element = self.browser_sei.find_element(By.XPATH, elemento_caixa_pesquisar)
        element.send_keys(self.numero_processo_sei)
   
        elemento_lupa = '//*[@id="spnInfraUnidade"]/img'
        element = self.browser_sei.find_element(By.XPATH, elemento_lupa)
        element.click()
 
        #espera o iframe estar presente        
        iframe = WebDriverWait(self.browser_sei, 120).until(
            EC.presence_of_element_located((By.ID, 'ifrVisualizacao'))
        )
 
        sleep(0.5)
 
        self.browser_sei.switch_to.default_content()
        self.browser_sei.switch_to.frame(iframe)
 
        elemento_incluir_documento = '//*[@id="divArvoreAcoes"]/a[1]/img'
        wait = WebDriverWait(self.browser_sei, 20)
        element = wait.until(EC.presence_of_element_located((By.XPATH, elemento_incluir_documento )))
        element.click()
       
        sleep(0.4)
 
        botao_declaracao = '//*[@id="tblSeries"]/tbody/tr[2]/td/a[2]'
        element = self.browser_sei.find_element(By.XPATH, botao_declaracao)
        element.click()
 
        sleep(0.4)
 
        botao_texto_padrao = '//*[@id="divOptTextoPadrao"]/div/label'
        element = self.browser_sei.find_element(By.XPATH, botao_texto_padrao)
        element.click()
 
        sleep(0.4)
 
        barra_pesquisa_declaracao = '//*[@id="txtTextoPadrao"]'
        element = self.browser_sei.find_element(By.XPATH, barra_pesquisa_declaracao)
        element.click()
 
        element.send_keys("DECLARAÇÃO VÍNCULO APOSENTADO")
       
        sleep(2)
 
        element.send_keys(Keys.ENTER)
 
        sleep(2)
 
        botao_restrito = '//*[@id="divOptRestrito"]/div/label'
        element = self.browser_sei.find_element(By.XPATH, botao_restrito)
        element.click()
 
        sleep(1)
 
        barra_procurar_hipotese_legal = '//*[@id="selHipoteseLegal"]'
        element = self.browser_sei.find_element(By.XPATH, barra_procurar_hipotese_legal)
        element.click()
 
        sleep(1)
 
        txt = 'Informação pessoal'
        element.send_keys(txt)
 
        sleep(1)
 
        element.send_keys(Keys.ENTER)
 
        opcao_info_pessoal = '//*[@id="selHipoteseLegal"]/option[9]'
        element = self.browser_sei.find_element(By.XPATH, opcao_info_pessoal)
        element.click()
 
        sleep(1)
 
        botao_salvar = '//*[@id="btnSalvar"]'
        element = self.browser_sei.find_element(By.XPATH, botao_salvar)
        element.click()
 
        self.browser_sei.switch_to.default_content()

        self.preencher_dados_declaracao(aposentado)       
       
        

    def preencher_dados_declaracao(self, aposentado):
        try:
            # Mude para a nova janela (a última na lista)
            all_windows = self.browser_sei.window_handles
            self.browser_sei.switch_to.window(all_windows[-1])
            print(f"entrou em {all_windows[-1]}")

            self.browser_sei.maximize_window()

            sleep(10)  # Aumentar o tempo de espera

            iframes = self.browser_sei.find_elements(By.TAG_NAME, 'iframe')
            print(f"Número de iframes: {len(iframes)}")
            for i, iframe in enumerate(iframes):
                print(f"Iframe {i}: {iframe.get_attribute('title')}")

            # XPaths dos elementos a serem editados
            iframe_tabela                   = '//iframe[@title="Corpo do Texto"]'
            elemento_nome                   = '/html/body/p[1]/span[2]/span/span/span/span/span/span/span/span/span/span'
            elemento_cpf                    = '/html/body/p[1]/span[4]/span/span/span/span/span/span/span/span/span/span/span'
            elemento_siape                  = '/html/body/p[1]/span[6]/span/span/span/span/span/span/span/span/span/span/span'
            elemento_orgao_origem           = '/html/body/p[1]/span[8]/span/span/span/span/span/span/span/span/span/span/span'
            elemento_data_aposentadoria     = '/html/body/p[1]/span[10]/span/span/span/span/span/span/span/span/span/span/span'
            elemento_fundamento_legal       = '/html/body/p[1]/span[13]/span/span/span/span/span/span/span/span/span/span/span'
            elemento_num_portaria           = '/html/body/p[1]/span[16]/span/span/span/span/span/span/span/span/span/span/span'
            elemento_data_dou               = '/html/body/p[1]/span[18]/span/span/span/span/span/span/span/span/span/span/span'

            
            # Alternar para o iframe pelo índice 2 (Corpo do texto)
            self.browser_sei.switch_to.frame(2)

            # Encontrar o elemento
            element = self.browser_sei.find_element(By.XPATH, elemento_nome)
            # Executar o script para alterar o texto
            self.browser_sei.execute_script("arguments[0].innerText = arguments[1];", element, aposentado.nome)

            sleep(0.5)

            element = self.browser_sei.find_element(By.XPATH, elemento_cpf)            
            self.browser_sei.execute_script("arguments[0].innerText = arguments[1];", element, aposentado.cpf)

            sleep(0.5)

            element = self.browser_sei.find_element(By.XPATH, elemento_siape)            
            self.browser_sei.execute_script("arguments[0].innerText = arguments[1];", element, aposentado.siape)

            sleep(0.5)

            element = self.browser_sei.find_element(By.XPATH, elemento_orgao_origem)            
            self.browser_sei.execute_script("arguments[0].innerText = arguments[1];", element, aposentado.orgao_origem)

            sleep(0.5)

            element = self.browser_sei.find_element(By.XPATH, elemento_data_aposentadoria)            
            self.browser_sei.execute_script("arguments[0].innerText = arguments[1];", element, aposentado.data_aposentadoria)

            sleep(0.5)

            element = self.browser_sei.find_element(By.XPATH, elemento_fundamento_legal)            
            self.browser_sei.execute_script("arguments[0].innerText = arguments[1];", element, aposentado.fundamento_legal)

            sleep(0.5)

            element = self.browser_sei.find_element(By.XPATH, elemento_num_portaria)            
            self.browser_sei.execute_script("arguments[0].innerText = arguments[1];", element, aposentado.dl_aposentadoria)

            sleep(0.5)

            element = self.browser_sei.find_element(By.XPATH, elemento_data_dou)            
            self.browser_sei.execute_script("arguments[0].innerText = arguments[1];", element, aposentado.data_dou)

            sleep(0.5)

            self.browser_sei.switch_to.default_content()

            sleep(5)

            # Salvar a tabela
            caminho_salvar_tabela = '//*[@id="cke_27"]'
            self.browser_sei.find_element(By.XPATH, caminho_salvar_tabela).click()
            
            sleep(2)            

            # Fechar a janela atual
            self.browser_sei.close()

            # Voltar para a janela principal (ou a janela original)
            self.browser_sei.switch_to.window(all_windows[0])

        except Exception as e:
            print(f"Ocorreu um erro durante o preenchimento dos dados do processo: {e}")


    def buscar_nome_orgao(self, caminho_arquivo, codigo_orgao):
        # Lê a planilha Excel
        df = pd.read_excel(caminho_arquivo)
        
        # Exibe os nomes das colunas para verificar se estão corretos
        print("Colunas encontradas na planilha:", df.columns)

        # Garante que o código do órgão seja tratado como string, caso esteja como número
        df['IT-CO-ORGAO'] = df['IT-CO-ORGAO'].astype(str)
        codigo_orgao = str(codigo_orgao)
        
        # Filtra a linha onde o código do órgão corresponde ao valor fornecido
        resultado = df[df['IT-CO-ORGAO'] == codigo_orgao]
        
        # Verifica se encontrou alguma correspondência
        if not resultado.empty:
            # Retorna o nome do órgão correspondente
            return resultado['IT-NO-ORGAO'].values[0]
        else:
            return "Órgão não encontrado"

    def grafo_objetos(self, lista_objetos):
        obj_itinerante = Vinculo

        # obj_itinerante é o objeto que viaja entre os demais objetos da lista setando cada um como visitado.
        
        '''
        O bloco abaixo testa se a lista é vazia, se tem só um objeto ou mais.
        Se é vazia retorna None, pois não se aplica a regra;
        Se tem somente um objeto de vínculo testa se é elegível para órgão de origem e retorna o objeto;
        Se tem mais de um varre a lista e procura pelo objeto obj_itinerante recebe o objeto "sem pai" para iniciar as visitações a partir dele.'''
        if len(lista_objetos) == 0:
            return None
                    
        elif len(lista_objetos) > 1:
            for obj in lista_objetos:
                if obj.orgao_anterior is None:
                    obj.is_visitado = True
                    obj.is_elegivel_orgao_origem = obj.is_regra_origem()
                    obj_itinerante = obj
                    

        else:
            for obj in lista_objetos:
                try:
                    obj.is_visitado = True
                    obj.is_elegivel_orgao_origem = obj.is_regra_origem()
                    return obj
                except Exception as erro:
                    print("-----------------------------------------")
                    print(erro)
                    print(f"Tamanho lista objeto: {len(lista_objetos)}")
                    print("-----------------------------------------")
                    msg("pause")
                    return None
        
        '''
        O bloco abaixo é um loop onde o obj_itinerante procura pelo órgão_atual. Quando encontra obj é setado como visitado,
        verifica-se se é elegível para órgão de origem e obj_itinerante recebe obj.
        A cada iteração verifica-se se o obj_itinerante possui órgão atual. Se sim, continua a busca. Se não sai do loop. Se obj_itinerante possui
        órgão atual e não encontra o obj com este órgão, sai do loop.
        '''
        continua_verificacao = True
        
        while continua_verificacao:

            for obj in lista_objetos:

                cond1 = obj_itinerante.orgao_atual is not None

                cond2 = False
                if obj_itinerante.orgao_atual is not None:
                    cond2 = obj_itinerante.orgao_atual.strip() == obj.orgao.strip() # strip() tira espaços no inicio e no final da string
                
                cond3 = obj.is_visitado == False

                #msg(f"cond1: {cond1}\ncond2: {cond2}\ncond3: {cond3}")

                if cond1 == False:
                    continua_verificacao = False
                    break

                if cond2 and cond3:
                    obj.is_visitado = True
                    obj.is_elegivel_orgao_origem = obj.is_regra_origem()
                    obj_itinerante = obj           

        '''
        Verifica se a lista de objetos foi toda setada como visitada. Se sim, procura-se pelo objeto elegível para órgão de origem e retorna-se
        este objeto. Se não, se nem todas foram visitadas, retorna-se None para que seja feita verificação manual.
        '''
        todos_visitados = True

        for obj in lista_objetos:  

            if obj.is_visitado:
                pass
            else:
                todos_visitados = False
                return None
        
        if todos_visitados:
            for obj in lista_objetos:
                if obj.is_elegivel_orgao_origem:
                    return obj                

    def rastrear_orgao_origem(self, lista_objetos):

        objeto_i = Vinculo()
        
        # Primeiro busca por regra e seleciona o vínculo de partida: objeto_i
        for obj in lista_objetos:
            
            if obj.is_regra_origem():
                objeto_i = obj
                objeto_i.is_elegivel_orgao_origem = True #Elege o vinculo como orgão de origem pois a ele se aplica regra. Ver def is_regra_origem em classe Vinculo.
                print("linha")
                print("VÍNCULO:\n")
                print(objeto_i)
            else:
                print("linha")
                print("vínculo:\n")
                print(obj)
                
        
        continua = True
        while continua:
            
            for obj in lista_objetos:

                if objeto_i.orgao_atual is not None:

                    if objeto_i.orgao_atual.strip() == obj.orgao.strip():

                        objeto_i.is_linkado = True #objeto itinerante é setado como linkado
                        objeto_i = obj #objeto itinerante recebe novo objeto
                        objeto_i.is_linkado = True

                        if objeto_i.is_regra_origem(): #testa se objeto itinerante obedece regra de origem e se sim é setado como elegível
                            objeto_i.is_elegivel_orgao_origem = True                        
                else:                    
                    continua = False
                    break

        flag_verifica_link = True
        for obj in lista_objetos:
            if obj.is_linkado:
                pass
            else:
                flag_verifica_link = False
        
        if flag_verifica_link:
            for obj in lista_objetos:
                if obj.is_elegivel_orgao_origem:                    
                    return obj
        else:
            return obj

    def coletar_dados_em_cacoaposse(self, aposentado):

        self.__app = Application().connect(title_re="^Terminal 3270.*")
        self.__dlg = self.__app.window(title_re="^Terminal 3270.*")
        self.__acesso_terminal = controle_terminal3270.Janela3270()

        self.navega_cacoaposse_ate_cpf()

        self.__dlg.type_keys(aposentado.cpf) #cpf
        sleep(0.5)
        self.__dlg.type_keys("{ENTER}")

        #Digita CPF + ENTER e verifica as possíveis situações abaixo

        texto1 = r'[_\s]*(RH NAO CADASTRADO)[_\s]*'
        texto2 = r'[_\s]*(NAO EXISTEM DADOS PARA ESTA CONSULTA)[_\s]*'
        texto3 = r'[_\s]*(VOCE NAO ESTA AUTORIZADO A CONSULTAR DADOS DESTE SERVIDOR)[_\s]*'

        texto_tela = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()

        tem_texto1 = re.findall(texto1, texto_tela)
        tem_texto2 = re.findall(texto2, texto_tela)
        tem_texto3 = re.findall(texto3, texto_tela)

        if tem_texto1:
            return "RH NAO CADASTRADO"
        elif tem_texto2:
            return "NAO EXISTEM DADOS PARA ESTA CONSULTA"
        elif tem_texto2:
            return "VOCE NAO ESTA AUTORIZADO A CONSULTAR DADOS DESTE SERVIDOR"
        else:
            pass

        sleep(0.5)
        self.__dlg.type_keys("x")
        sleep(0.5)
        self.__dlg.type_keys("{ENTER}")
        sleep(0.5)

        flag_avanca_pagina          = True
        flag_texto_dl_aposentadoria = False
        flag_texto_data_inicio      = False
        flag_texto_fundamento_legal = False

        while flag_avanca_pagina:

            for i in range(1, 25):

                if flag_texto_fundamento_legal and flag_texto_data_inicio and flag_texto_dl_aposentadoria:
                    flag_avanca_pagina = False
                    break

                texto_ultima_pagina         = r'[_\s]*(ULTIMA TELA DO VINCULO)[_\s]*'
                texto_f8_avanca             = r'[_\s]*(PF8=AVANCA)[_\s]*'
                texto_dl_aposentadoria      = r'[_\s]*(DL APOSENTADORIA)[_\s]*'
                texto_data_inicio           = r'[_\s]*(DATA INICIO)[_\s]*'
                texto_fundamento_legal      = r'[_\s]*(FUNDAMENTO LEGAL)[_\s]*'
                

                texto_tela = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()
                texto_linha = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), i, 1, i, 80).strip()

                tem_texto_dl_aposentadoria = re.findall(texto_dl_aposentadoria, texto_linha)
                if tem_texto_dl_aposentadoria:
                    flag_texto_dl_aposentadoria = True
                    aposentado.dl_aposentadoria = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), i, 20, i, 80).strip()
                    

                    # Expressão regular para datas abreviadas (ddMMMyyyy)
                    padrao_abreviado = r'(\d{2})(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)(\d{4})'

                    # Expressão regular para datas por extenso (dd de mês de aaaa)
                    padrao_extenso = r'(\d{2})\sde\s(janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\sde\s(\d{4})'

                    # Expressão regular para capturar o padrão "DO. dd.mm.yy"
                    padrao_do = r'DO\.\s(\d{2}\.\d{2}\.\d{2})'

                    # Primeiro tenta encontrar a data no formato abreviado
                    match_abreviado = re.search(padrao_abreviado, aposentado.dl_aposentadoria)

                    match_extenso = re.search(padrao_extenso, aposentado.dl_aposentadoria)

                    match_do = re.search(padrao_do, aposentado.dl_aposentadoria)

                    if match_abreviado:
                        aposentado.data_dou = self.converter_data(match_abreviado)  # Funciona como antes
                    elif match_extenso:
                        aposentado.data_dou = self.converter_data(match_extenso)  # Funciona como antes
                    elif match_do:
                        # Para o caso DO. dd.mm.yy, criamos manualmente o "match" esperado
                        dia, mes, ano_abreviado = match_do.group(1).split('.')

                        # Expande o ano abreviado para quatro dígitos
                        ano = f"19{ano_abreviado}" if int(ano_abreviado) > 50 else f"20{ano_abreviado}"

                        # Como o mês está no formato numérico (e não abreviado), fazemos a conversão direta
                        data_formatada = f"{dia}/{mes}/{ano}"

                       
                        # Passamos a data formatada diretamente, sem usar converter_data, pois já formatamos corretamente
                        aposentado.data_dou = data_formatada
                        
                       
                    else:
                        aposentado.data_dou = "Não foi possível encontrar a Data DOU."
                        aposentado.status = "Verificação manual: Data DOU não encontrada."
                        

                    

                tem_texto_data_inicio = re.findall(texto_data_inicio, texto_linha)
                if tem_texto_data_inicio:
                    flag_texto_data_inicio = True
                    texto = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), i, 1, i, 30).strip()
                    padrao_data = r'DATA INICIO     : (\d{2})(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)(\d{4})'
                    match = re.search(padrao_data, texto)
                    aposentado.set_data_aposentadoria(self.converter_data(match))
                    

                tem_texto_fundamento_legal = re.findall(texto_fundamento_legal, texto_linha)
                if tem_texto_fundamento_legal:
                    flag_texto_dl_aposentadoria = True
                    aposentado.fundamento_legal = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), i, 24, i, 80).strip()

                if i >= 24: # Só procura ultima página se linha for maior igual a 24. Assim garante que a última página seja lida até a última linha.
                    
                        tem_texto_f8_avanca = re.findall(texto_f8_avanca, texto_tela)
                        
                        if tem_texto_f8_avanca:

                            tem_ultima_pagina = re.findall(texto_ultima_pagina, texto_tela)
                            if tem_ultima_pagina:
                                while True:
                                    texto_tela = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()
                                    texto_opcoes = r'[_\s]*(INFORME UMA DAS OPCOES)[_\s]*'
                                    if texto_opcoes:
                                        break
                                    self.__dlg.type_keys('{F12}')
                                    sleep(1)
                                sleep(1)
                                self.__dlg.type_keys('{TAB}')
                                flag_avanca_pagina = False
                                break
                        else:
                            while True:
                                texto_tela = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()
                                texto_opcoes = r'[_\s]*(INFORME UMA DAS OPCOES)[_\s]*'
                                if texto_opcoes:
                                    break
                                self.__dlg.type_keys('{F12}')
                                sleep(1)
                                
                            sleep(1)
                            self.__dlg.type_keys('{TAB}')
                            flag_avanca_pagina = False
                            break

            if flag_avanca_pagina:
                    self.__dlg.type_keys('{F8}')

        return "OK"

    
    def coletar_dados_em_cdconvinc(self, lista_tuplas):
        lista_vinculos = [] 

        self.__app = Application().connect(title_re="^Terminal 3270.*")
        self.__dlg = self.__app.window(title_re="^Terminal 3270.*")
        self.__acesso_terminal = controle_terminal3270.Janela3270()
        

        sleep(0.5)
        for tupla in lista_tuplas: 

            self.navega_cdconvinc_ate_cpf()

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

            vinculo = Vinculo()
            vinculo.orgao = tupla[5]

            while flag_avanca_pagina:

                for i in range(1, 25):
                    texto_aposentadoria     = r'[_\s]*(APOSENTADORIA)[_\s]*'
                    texto_ingresso          = r'[_\s]*(INGRESSO NO ORGAO)[_\s]*'
                    texto_exclusao          = r'[_\s]*(EXCLUSAO)[_\s]*'
                    texto_redistribuicao    = r'[_\s]*(REDISTRIBUICAO / REFORMA)[_\s]*'
                    texto_ultima_pagina     = r'[_\s]*(ULTIMA TELA DO VINCULO)[_\s]*'
                    texto_f8_avanca         = r'[_\s]*(PF8=AVANCA)[_\s]*'
                    texto_f8_avanca         = r'[_\s]*(PF8=AVANCA)[_\s]*'

                    texto_tela = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()
                    texto_linha = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), i, 1, i, 80).strip()
                    msg_log = "Linha "+str(i)+": "+texto_linha
                    #print(msg_log)
                    
                                        
                    tem_aposentadoria = re.findall(texto_aposentadoria, texto_linha)
                    if tem_aposentadoria:
                        #msg("tem aposentadoria")
                        print("TEM APOSENTADORIA")
                        data = None
                        for j in range(i, i+3):                                                    
                            texto = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), j, 1, j, 80).strip()
                            print(texto)
                            padrao_data = r'DATA OCORRENCIA:(\d{2})(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)(\d{4})'
                            match = re.search(padrao_data, texto)
                            data = self.converter_data(match)
                            print("Data aposentadoria tratada: "+str(data))
                            vinculo.aposentadoria = data
                            if data is not None:
                                break

                    tem_ingresso = re.findall(texto_ingresso, texto_linha)
                    if tem_ingresso:
                        #msg("tem ingresso")
                        print("TEM INGRESSO")
                        data = None
                        for j in range(i, i+3):                                                    
                            texto = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), j, 1, j, 80).strip()
                            print(texto)
                            padrao_data = r'DATA OCORRENCIA:(\d{2})(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)(\d{4})'
                            match = re.search(padrao_data, texto)
                            data = self.converter_data(match)
                            print("Data ingresso tratada: "+str(data))
                            vinculo.ingresso = data
                            if data is not None:
                                break
                    
                    tem_exclusao = re.findall(texto_exclusao, texto_linha)
                    if tem_exclusao:
                        linha_de_baixo = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), (i+1), 1, (i+1), 80).strip()
                        #msg("tem exclusão")
                        texto_grupo_cocorrencia = r'[_\s]*(GRUPO/OCORRENCIA)[_\s]*'
                        tem_grupo_cocorrencia = re.findall(texto_grupo_cocorrencia, linha_de_baixo)

                        if tem_grupo_cocorrencia:
                            print("TEM EXCLUSAO / TEM GRUPO OSORRÊNCIA")
                        
                            #msg("tem grupo/ocorrencia")
                            data = None
                            for j in range(i+2, i+3):                                                    
                                texto = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), j, 1, j, 80).strip()
                                print(texto)
                                padrao_data = r'DATA OCORRENCIA:(\d{2})(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)(\d{4})'
                                match = re.search(padrao_data, texto)
                                data = self.converter_data(match)
                                vinculo.exclusao = data
                                print("Data exclusao tratada: "+str(data))
                                if data is not None:
                                    break
                    
                    tem_redistribuicao = re.findall(texto_redistribuicao, texto_linha)
                    if tem_redistribuicao:
                        #msg(" redistribuição")

                        texto = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), (i+1), 27, (i+1), 41).strip()
                        print(f"Redistribuição matriz: {texto}")
                        
                        # Coleta órgao e matrícula anteriores
                        if texto and texto is not None and texto != "":
                            partes_entrada  = texto.split("/")
                            orgao           = partes_entrada[0]
                            matricula       = partes_entrada[1]
                            vinculo.orgao_anterior = orgao
                            print(f"Órgão matriz: {orgao}")
                            vinculo.matricula_anterior = matricula
                            print(f"Matrícula matriz: {matricula}")
                        else:
                            vinculo.orgao_anterior = None
                            vinculo.matricula_anterior = None

                        # Coleta órgao e matrícula atuais
                        texto = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), (i+2), 27, (i+2), 41).strip()
                        if texto and texto is not None and texto != "":
                            partes_entrada  = texto.split("/")
                            orgao           = partes_entrada[0]
                            matricula       = partes_entrada[1]
                            vinculo.orgao_atual = orgao
                            print(f"Órgão matriz: {orgao}")
                            vinculo.matricula_atual = matricula
                            print(f"Matrícula matriz: {matricula}")
                        else:
                            vinculo.orgao_atual = None
                            vinculo.matricula_atual = None
                        
                    if i >= 24: # Só procura ultima página se linha for maior igual a 24. Assim garante que a última página seja lida até a última linha.
                    
                        tem_texto_f8_avanca = re.findall(texto_f8_avanca, texto_tela)
                        
                        if tem_texto_f8_avanca:

                            tem_ultima_pagina = re.findall(texto_ultima_pagina, texto_tela)
                            if tem_ultima_pagina:
                                while True:
                                    texto_tela = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()
                                    texto_opcoes = r'[_\s]*(INFORME UMA DAS OPCOES)[_\s]*'
                                    if texto_opcoes:
                                        break
                                    self.__dlg.type_keys('{F12}')
                                    sleep(1)
                                sleep(1)
                                self.__dlg.type_keys('{TAB}')
                                flag_avanca_pagina = False
                                break
                        else:
                            while True:
                                texto_tela = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()
                                texto_opcoes = r'[_\s]*(INFORME UMA DAS OPCOES)[_\s]*'
                                if texto_opcoes:
                                    break
                                self.__dlg.type_keys('{F12}')
                                sleep(1)
                                
                            sleep(1)
                            self.__dlg.type_keys('{TAB}')
                            flag_avanca_pagina = False
                            break                   
                    
                if flag_avanca_pagina:
                    self.__dlg.type_keys('{F8}')                

            lista_vinculos.append(vinculo) 

        return lista_vinculos        
        
    def navega_cdconvinc_ate_cpf(self):
        sleep(0.5)
        self.__dlg.type_keys('{F3}')
        sleep(0.5)
        self.__dlg.type_keys('{TAB 9}')
        sleep(0.5)                    
        self.__dlg.type_keys('>'+'CDCONVINC')
        sleep(0.5)
        self.__dlg.type_keys("{ENTER}")                  
        sleep(0.5)
        self.__dlg.type_keys('{TAB}')

    def navega_cacoaposse_ate_cpf(self):
        sleep(0.5)
        self.__dlg.type_keys('{F3}')
        sleep(0.5)
        self.__dlg.type_keys('{TAB 9}')
        sleep(0.5)                    
        self.__dlg.type_keys('>'+'CACOAPOSSE')
        sleep(0.5)
        self.__dlg.type_keys("{ENTER}")                  
        sleep(0.5)
        self.__dlg.type_keys('{TAB 4}')
        sleep(0.5)
        self.__dlg.type_keys('x')    
    
                    
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
            
    def converter_data(self, match):       
        
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
        colunas_necessarias = ['Status', 'CPF', 'Nome', 'SIAPE', 'Vínculo Decipex', 'Órgão de Origem', 'Data Aposentadoria', 'Fundamento Legal', 'Portaria Número', 'Data Publicação DOU']
        
        for coluna in colunas_necessarias:
            if coluna not in base_dados_atualizada.columns:
                base_dados_atualizada[coluna] = None

        # Reorganiza as colunas
        colunas_organizadas = colunas_necessarias + [coluna for coluna in base_dados_atualizada.columns if coluna not in colunas_necessarias]
        base_dados_atualizada = base_dados_atualizada[colunas_organizadas]

        # Atualiza os dados das respectivas colunas para cada aposentado na lista
        for aposentado in lista_aposentados:
            linha = aposentado.linha_planilha
            base_dados_atualizada.at[linha, 'Status']               = aposentado.status
            base_dados_atualizada.at[linha, 'CPF']                  = str(aposentado.cpf)
            base_dados_atualizada.at[linha, 'Nome']                 = aposentado.nome           
            base_dados_atualizada.at[linha, 'Vínculo Decipex']      = aposentado.vinculo_decipex
            base_dados_atualizada.at[linha, 'SIAPE']                = aposentado.siape
            base_dados_atualizada.at[linha, 'Órgão de Origem']      = aposentado.orgao_origem
            base_dados_atualizada.at[linha, 'Data Aposentadoria']   = aposentado.data_aposentadoria
            base_dados_atualizada.at[linha, 'Fundamento Legal']     = aposentado.fundamento_legal
            base_dados_atualizada.at[linha, 'Portaria Número']      = aposentado.dl_aposentadoria
            base_dados_atualizada.at[linha, 'Data Publicação DOU']  = aposentado.data_dou

        # Salva a planilha Excel atualizada, convertendo todos os dados para string
        base_dados_atualizada = base_dados_atualizada.astype(str)
        base_dados_atualizada.to_excel(self.planilha, index=False)

    def log(self,mensagem):
        """
        Adiciona uma mensagem ao arquivo de log com a data e hora atuais.

        Parâmetros:
        - mensagem (str): A mensagem que será adicionada ao log.
        - nome_arquivo (str): O nome do arquivo de log. Padrão é 'log.txt'.
        """
        # Obtém a data e hora atual
        agora = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
        
        # Cria ou abre o arquivo de log no modo de adição
        with open('log.txt', 'a', encoding='utf-8') as arquivo_log:
            # Escreve a mensagem com a data e hora
            if mensagem == "linha":
                arquivo_log.write(f'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_\n')
            else:
                #arquivo_log.write(f'[{agora}] {mensagem}\n')
                arquivo_log.write(f'{mensagem}\n')

if __name__ == "__main__":
        bot = Decaposi()
        sys.exit()