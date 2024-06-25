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
    def __init__(self):
        pass

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
        titulo_texto = r'[_\s]*(INFORME UMA DAS OPCOES)[_\s]*'
        nome_texto = r'[_\s]*(NOME)[_\s]*'
        cpf_texto = r'[_\s]*(CPF)[_\s]*'

        resultado1 = re.findall(titulo_texto, texto_tela)
        resultado2 = re.findall(nome_texto, texto_tela)
        resultado3 = re.findall(cpf_texto, texto_tela)
        
        if resultado1 and resultado2 and resultado3:
            return True
        else:
            return False