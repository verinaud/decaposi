from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from pywinauto import Application
from selenium import webdriver
import controle_terminal3270
from time import sleep
import keyboard as kb
import glob
import sys
import re
import os



class CACOAPOSSE:
    def __init__(self, url):
        self.__url                          = url
        self.__acesso_terminal              = None

        self.__lista_tuplas                 = []
        self.__lista_cpf_ja_consultados     = []


    def iniciar_cacoaposse(self):
        """
        Descrição:
            Inicia Terminal 3270 e navega até Cacoaposse no campo CPF.

        Parâmetros:
            Nenhum.

        Retorno:
            Nenhum.
        
        """
        self.__sair_siape()

        self.__base_path = os.path.dirname(os.path.abspath(__file__))

        self.__download_folder = os.path.join(self.__base_path, 'Downloads')

        # Criar o diretório 'downloads' se não existir
        if not os.path.exists(self.__download_folder):
            os.makedirs(self.__download_folder)

        # Limpar a pasta de downloads
        for arquivo in glob.glob(os.path.join(self.__download_folder, 'hodcivws*.jsp')):
            os.remove(arquivo)

        options = Options()
        prefs = {
            "download.default_directory": self.__download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "credentials_enable_service": False,
            "password_manager_enabled": False
        }
        # Removido o caminho binário do Chrome, já que vamos usar o brwoser instalado
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_argument('disable-infobars')

        # Usando ChromeDriverManager para gerenciar o driver
        brwoser = webdriver.Chrome(options=options)

        brwoser.maximize_window()

        brwoser.get(self.__url)
        wait = WebDriverWait(brwoser, 10)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="linkCD"]/img'))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menu"]/ul[3]/li[1]/a/span'))).click()

        try:
            details_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="details-button"]')))
            if details_button:
                details_button.click()
                wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="proceed-link"]'))).click()
        except Exception:
            pass  # Detalhes ou botão de prosseguir não encontrados

        flag = True
        conta_flag = 0
        while flag :
            try:
                #sleep(3)  # Pode ser ajustado conforme a conexão
                __download_folder = self.__download_folder  # Acessando o caminho da pasta de downloads
                latest_file = max(glob.glob(os.path.join(__download_folder, 'hodcivws*.jsp')), key=os.path.getctime)
                brwoser.close()
                os.system(f'"{latest_file}"')  # Executa o arquivo baixado
                flag = False
            except Exception:
                sleep(1)
                conta_flag+=1
                if conta_flag >= 10 :
                    flag = False

        #se aparecer a tela de Advertência de Segurança, fecha...
        flag = True
        conta_flag = 0
        while flag :
            try :                
                __app = Application().connect(title_re="^Advertência de Segurança.*")
                __dlg = __app.window(title_re="^Advertência de Segurança.*")
                __dlg.type_keys("{TAB 2}")
                __dlg.type_keys('{ENTER}')
                flag = False
            except Exception as e :
                sleep(1)
                conta_flag+=1
                if conta_flag >= 10 :
                    flag = False

        #O bloco abaixo aguardar o Terminal 3270 abrir por 30 segundos, retornando True ou False
        flag = True
        conta_flag = 0
        while flag :
            try :
                self.__app = Application().connect(title_re="^Terminal 3270.*")
                self.__dlg = self.__app.window(title_re="^Terminal 3270.*")
                self.__acesso_terminal = controle_terminal3270.Janela3270()
                flag = False
                
            except Exception as e :
                sleep(1)
                conta_flag+=1
                if conta_flag >= 60 :
                    flag = False
                    return False
                
        flag = True
        conta_flag = 15
                
        while flag :
            conta_flag-=1
            print("Tentando abrir Terminal 3270:",conta_flag)
            texto_tela = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()
            padrao = r'[_\s]*(SIAPE)[_\s]*'
            resultado = re.findall(padrao, texto_tela)
            if resultado:
                flag = False
                sleep(0.3)
            if conta_flag == 0:
                self.__sair_siape()
                return False
            sleep(1)

        self.__navegacao_inicial_ate_cpf()

        sleep(1)

        texto_tela = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()
        titulo_texto = r'[_\s]*(SELECIONE UMA DAS OPCOES ABAIXO)[_\s]*'
        cpf_texto = r'[_\s]*(APOSENTADORIAS DO CPF)[_\s]*'

        resultado1 = re.findall(titulo_texto, texto_tela)
        resultado2 = re.findall(cpf_texto, texto_tela)
        
        if resultado1 and resultado2:
            return True
        else:
            return False
        
    def __navegacao_inicial_ate_cpf(self):
        
        texto_tela = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()
        texto1 = r'[_\s]*(MENSAGEM)[_\s]*'
        texto2 = r'[_\s]*(CADASTRADA)[_\s]*'
        texto3 = r'[_\s]*(POSICIONE O CURSOR NA OPCAO DESEJADA E PRESSIONE)[_\s]*'
        
        resultado1 = re.findall(texto1, texto_tela)
        resultado2 = re.findall(texto2, texto_tela)
        resultado3 = re.findall(texto3, texto_tela)

        if resultado1 and resultado2 and  resultado3:
            sleep(0.5)
            self.__dlg.type_keys('{F3}')
            self.__dlg.type_keys('{F2}')                   
            sleep(0.1)                    
            self.__dlg.type_keys(">"+'CACOAPOSSE')
            sleep(0.1)
            kb.press("Enter")                    
            sleep(0.1)
            self.__dlg.type_keys('{TAB}')
                        
        else:
            sleep(0.5)
            self.__dlg.type_keys('{F3}')
            self.__dlg.type_keys('{F2}')                   
            sleep(0.1)                    
            self.__dlg.type_keys(">"+'CACOAPOSSE')
            sleep(0.1)
            kb.press("Enter")
            self.__dlg.type_keys('{TAB}')     
    def __sair_siape(self) :
        try:
            __app = Application().connect(title_re="^Painel de controle.*")
            __dlg = __app.window(title_re="^Painel de controle.*")
            __dlg.type_keys('%{F4}')            
            flag = True
            sleep(0.5)

            while flag :

                try :
                    __app = Application().connect(title_re="^Fechar.*")
                    __dlg = __app.window(title_re="^Fechar.*")
                    # Pressionando ENTER
                    __dlg.type_keys('{ENTER}')
                    flag = False

                except Exception :
                    sleep(1)
                    conta_flag+=1
                    if conta_flag >= 15 :
                        flag = False            

        except Exception:
            pass
    def cpf_ja_consultado(self, cpf):
        flag = False
        for x in self.__lista_cpf_ja_consultados:           
            
            if str(cpf) == str(x):
                print("----------------------------------")
                print(cpf)
                print(x)
                print("----------------------------------")
                flag = False
                break

        return flag
    
    def get_lista_dados(self, cpf):
        """
        Descrição:
            Fornece uma lista de tuplas com dados referente ao cpf informado. Cada tupla da lista contém
            respectivamente:
                Número do CPF - string;
                Vínculo Decipex - boolean;
                Serv/Pens - string;
                Página - integer;
                Linha - integer;
                Órgão - string;
                Nota - string.

        Parâmetros:
            CPF - string.

        Retorno:
            lista - list.
        
        """
        lista = []
        for x in self.__lista_tuplas:
            if x[0] == cpf:
                lista.append(x)
        return lista

    def get_vinculo_decipex(self, cpf):
        """
        Descrição:
            Informa se CPF consultado tem vínculo com o Decipex.

        Parâmetros:
            CPF - string.

        Retorno:
            True/False - bool.
        
        """
        flag = False
        for x in self.__lista_tuplas:
            if x[0] == cpf and x[1]:
                flag = True
        return flag
    
    def get_status_cpf(self, cpf):
        """
        Descrição:
            Informa o status do CPF.

        Parâmetros:
            CPF - string.

        Retorno:
            Mensagem de nota - string.
        
        """
        status = ""
        for x in self.__lista_tuplas:
            if x[0] == cpf:
                status = x[6]
        return status

    def consultar_cpf(self, cpf):
        """
        Descrição:
            Consulta o CPF em CACOAPOSSE.

        Parâmetros:
            CPF - string.

        Retorno:
            Sem retorno.
        
        """
        if self.cpf_ja_consultado(cpf):
            pass
        else:
            self.__consultar_cpf(cpf)
    
    def __consultar_cpf(self, cpf):
        
        self.__dlg.type_keys("{TAB 3}")
        sleep(0.5)
        self.__dlg.type_keys("x")
        sleep(0.5)

        self.__dlg.type_keys(cpf)
        sleep(1)
        self.__dlg.type_keys('{ENTER}')
        

        flag_prossiga = True
        conta_flag_prossiga = 0
        max_tentativas = 30  # Limite de tentativas
        conta_tentativa = 0

        while flag_prossiga:
            conta_flag_prossiga+=1

            texto_consulta = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()
            texto1 = r'[_\s]*(RH NAO CADASTRADO)[_\s]*'
            texto2 = r'[_\s]*(INFORME APENAS)[_\s]*'
            texto3 = r'[_\s]*(DIGITO VERIFICADOR INVALIDO)[_\s]*'
            texto4 = r'[_\s]*(NAO EXISTEM DADOS PARA ESTA CONSULTA)[_\s]*'
            texto5 = r'[_\s]*(MATRICULA NOME DO APOSENTADO)[_\s]*'
            texto6 = r'[_\s]*(VOCE NAO ESTA AUTORIZADO A CONSULTAR DADOS DESTE SERVIDOR)[_\s]*'

            resultado1 = re.findall(texto1, texto_consulta)
            resultado2 = re.findall(texto2, texto_consulta)
            resultado3 = re.findall(texto3, texto_consulta)
            resultado4 = re.findall(texto4, texto_consulta)
            resultado5 = re.findall(texto5, texto_consulta)
            resultado6 = re.findall(texto6, texto_consulta)

            if conta_flag_prossiga >= max_tentativas:
                print("Número máximo de tentativas atingido.")
                flag_prossiga = False
                return False

            if  resultado1 or resultado2 or resultado3 or resultado4 or resultado6:
                print(f"O número de CPF {cpf} não é válido! Verifique.")
                self.__lista_cpf_ja_consultados.append(cpf)
                flag_prossiga = False
                self.__mensagem_erro = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 24, 1, 24, 80).strip()

                sleep(0.5)
                self.__dlg.type_keys('{F3}')
                self.__dlg.type_keys('{F2}')                   
                sleep(0.1)                    
                self.__dlg.type_keys(">"+'CACOAPOSSE')
                sleep(0.1)
                kb.press("Enter")
                self.__dlg.type_keys('{TAB}') 
                continue

            if resultado5:
                self.__dlg.type_keys('x')
                sleep(0.2)
                self.__dlg.type_keys('{ENTER}')
                flag_selecione = True
                conta_flag_selecione = 0
                flag_prossiga = False

                while flag_selecione:
                    conta_flag_selecione+=1
                    texto_consulta1 = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()
                    texto11 = r'[_\s]*(DADOS DE ENTRADA NA APOSENTADORIA)[_\s]*'
                    texto12 = r'[_\s]*(DL APOSENTADORIA)[_\s]*'
                    texto13 = r'[_\s]*(DATA INICIO)[_\s]*'

                    resultado11 = re.findall(texto11, texto_consulta1)
                    resultado12 = re.findall(texto12, texto_consulta1)
                    resultado13 = re.findall(texto13, texto_consulta1)

                    if resultado11 and resultado12 and resultado13:            
                        flag_selecione = False
                        dl_aposentadoria    = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 10, 1, 10, 80).strip()
                        data_inicio         = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 11, 1, 11, 33).strip()  

                        sleep(0.5)                             
                        self.__dlg.type_keys('{F8 2}')
                        sleep(0.5)
                        
                        while True:
                            # Captura o conteúdo da tela
                            tela = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 1, 1, 24, 80).strip()

                            # Busca informações do fundamento legal
                            texto_consulta2 = self.__acesso_terminal.pega_texto_siape(tela, 1, 1, 24, 80).strip()
                            texto = r'[_\s]*(DADOS DO FUNDAMENTO LEGAL)[_\s]*'
                            texto0 = r'[_\s]*(FUNDAMENTO LEGAL)[_\s]*'

                            resultado9 = re.findall(texto, texto_consulta2)
                            resultado8 = re.findall(texto0, texto_consulta2)

                            if resultado9 and resultado8:
                                fundamento_legal = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 12, 2, 12, 80).strip()
                                self.__popula_tupla(cpf, dl_aposentadoria, data_inicio, fundamento_legal)
                                return  # Para parar após encontrar o fundamento legal

            else:
                conta_tentativa += 1
                sleep(0.5)
                if conta_tentativa > 6:
                    mensagem_erro = self.__acesso_terminal.pega_texto_siape(self.__acesso_terminal.copia_tela(), 24, 1, 24, 80).strip()
                    print(f"Deu erro {mensagem_erro}")
                    sleep(0.5)
                    self.__dlg.type_keys('{F3}')
                    self.__dlg.type_keys('{F2}')                   
                    sleep(0.1)                    
                    self.__dlg.type_keys(">"+'CACOAPOSSE')
                    sleep(0.1)
                    kb.press("Enter")
                    self.__dlg.type_keys('{TAB}') 
                    break 



    def __popula_tupla(self, cpf, dl_aposentadoria, data_inicio, fundamento_legal):
        tupla = (cpf, dl_aposentadoria, data_inicio, fundamento_legal)
        self.__lista_tuplas.append(tupla)
        print(tupla)
        

if __name__ == "__main__":
        lista_cpf = {"06194362715","00214850030","02857294700","00467931003","37094394004","06194362715","44043716753","37739891215","27780120104","14910195060","35511680753","14910195068","00455865515","05152470810","08749872885","50375725687","27780120104","62750887704","14910195068","25393430744","02058766768","21966311915","45883076734","02214792520","57050759872","42544173653","00063339315","20903073749","91149533820","09154388600","33810150606","72758007720","07007591744","01447849710","23830921772","14898217249","88171558704","04883063372","25597027468","72758007720","10318950782","02089440872","53457749604","02605708004","32048114768","50958631891","00653187491","02544967846","20565429272","15212971420","00028894200","56013884749","36342858772","05291616806","89221320278","14970031304","18694438887","35753528953","19572190687","01904944760","21966311915","02058766768","25393430744","00063339315","00063339315","02214792520","45883076734","99497948204","24467499687","27179192015","11104660563","36172128833","85117811704","21902836472","01220106984","14238152387","79956513504","01790664802","03919978315","68704690753","49674765700","03845338253","44420226749","07077809773","21707863687","39497798172","89925742587","02827603500","02858002800","47537744734","10914196472","35516410649","51879840634","29178770068","18232388234","17562210225","42449227687","32421281687","35817674734","31617182753","31644740672","15147002391","76075915834","89215575715","73952800759","21757917349","55760139649","75436019404","24652296720","25067311991","62487019700","02532840404","22065300744","12580279768","79113427849","05780322104"        }
        print(len(lista_cpf))
        sleep(5)

        url_siapenet = "https://www1.siapenet.gov.br/orgao/Login.do?method=inicio"

        cd = CACOAPOSSE(url_siapenet)

        cd.get_lista_dados

        cd.iniciar_cacoaposse()

        for cpf in lista_cpf:

            cd.consultar_cpf(cpf)           

            lista = cd.get_lista_dados(cpf)

            for x in lista:
                print(x)
                
            print(f"Tem dados disponiveis {cpf}?")

        #cd.consultar_cpf("00653187491")
                    
        sys.exit()