import os
import time
import glob
import keyboard as kb
from tela_mensagem import Mensagem as msg
from selenium.webdriver.support.ui import WebDriverWait
from selenium_setup import SeleniumSetup
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError
#from ControleTerminal3270 import JanelaPrint


# Inicia a abertura do módulo SIAPE pelo navegador
class IniciaModSiape:
    def __init__(self, siape, url):
        self.siape      = siape
        self.navegador  = siape.setup_driver()
        self.url        = url

    def navigate_to_site_and_click(self):
        self.navegador.get(self.url)
        wait = WebDriverWait(self.navegador, 10)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="linkCD"]/img'))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menu"]/ul[3]/li[1]/a/span'))).click()
                                                          #//*[@id="menu"]/ul[3]/li[1]/a/span       

        try:
            details_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="details-button"]')))
            if details_button:
                details_button.click()
                wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="proceed-link"]'))).click()
        except Exception:
            pass  # Detalhes ou botão de prosseguir não encontrados

    def _download_and_execute_file(self):
        time.sleep(3)  # Pode ser ajustado conforme a conexão
        download_folder = self.siape.download_folder  # Acessando o caminho da pasta de downloads
        latest_file = max(glob.glob(os.path.join(download_folder, 'hodcivws*.jsp')), key=os.path.getctime)
        self.navegador.close()
        os.system(f'"{latest_file}"')  # Executa o arquivo baixado
        

    def _connect_to_app(self, title, max_attempts=50):
        attempts = 0
        while attempts < max_attempts:
            try:
                return Application().connect(title=title)
            except Exception as e:
                print(f"Tentativa {attempts + 1}: {str(e)}")
                time.sleep(1)
                attempts += 1
        print(f"Não foi possível estabelecer uma conexão com {title}")
        return None

    def _interact_with_control_panel(self, app):
        dlg = app['Painel de controle']
        dlg.set_focus()
        time.sleep(2)
        for _ in range(4):
            dlg.type_keys('{TAB}')
            time.sleep(0.4)
        dlg.type_keys('{ENTER}')

    def _save_pid_of_window(self, title_pattern):
        app = Application().connect(title_re=title_pattern)
        try:
            app_conn = app.connect(title_re=title_pattern)
            pid = app_conn.process
            with open('pid.txt', 'w') as f:
                f.write(str(pid))
            return True
        except ElementNotFoundError:
            print("A janela não está presente.")
            return False

    def executar_siape(self):
        #driver = self.navegador.setup_driver()
        self.navigate_to_site_and_click()
        self._download_and_execute_file()

        #se aparecer a tela de Advertência de Segurança, fecha...
        flag = True
        conta_flag = 0
        while flag :
            try :                
                self.app = Application().connect(title_re="^Advertência de Segurança.*")
                self.dlg = self.app.window(title_re="^Advertência de Segurança.*")
                self.dlg.type_keys("{TAB 2}")
                kb.press("Enter")
                flag = False
            except Exception as e :
                time.sleep(1)
                conta_flag+=1
                if conta_flag >= 15 :
                    flag = False
    


class Terminal3270Connection:
    def __init__(self):
        self.app = None
        self.dlg = None
        self.connect_to_terminal()

    def connect_to_terminal(self):
        try:
            self.app = Application().connect(title_re="^Terminal 3270.*")
            self.dlg = self.app.window(title_re="^Terminal 3270.*")
        except Exception as e:
            print(f"Erro ao conectar ao Terminal 3270: {e}")
            # Tratamento de exceções ou reconexão pode ser feito aqui

    def mantem_hod_ativo(self):
        if not self.dlg:
            print("Não conectado ao Terminal 3270.")
            return
        try:
            self.dlg.type_keys('{F2}')
        except Exception as e:
            print(f"Erro ao enviar comandos para o Terminal 3270: {e}")


class ImprimePdfSiape:
    def __init__(self, nomearquivo):
        self.nomearquivo = nomearquivo

    def imprimir_pdf(self):
        # Localizar a janela pelo título
        app = Application().connect(title_re="^Terminal 3270.*")

        # Obter a janela
        window = app.window(title_re="^Terminal 3270.*")

        # Ativar a janela
        window.set_focus()
        # Pressionar as teclas de seta para navegar até o menu "Arquivo"
        kb.press('ctrl')
        kb.press('p')
        kb.release('p')
        kb.release('ctrl')

        time.sleep(3)

        app = Application().connect(title_re="Imprimir")
        dlg = app['Imprimir']
        time.sleep(1)

        # Obter a janela
        window = app.window(title_re="Imprimir")

        # Ativar a janela
        window.set_focus()
        dlg.type_keys('Microsoft Print To PDF')
        time.sleep(0.5)
        kb.press("Enter")

        app = Application().connect(title=u'Salvar Saída de Impressão como', timeout=5)
        dlg = app[u'Salvar Saída de Impressão como']

        arquiv = self.nomearquivo + ".pdf"
        arquiv = arquiv.replace(" ", "{SPACE}")
        caminho_script_atual = os.path.abspath(os.path.dirname(__file__))
        pasta = os.path.join(caminho_script_atual, 'pdfs')

        caminho = os.path.join(pasta, arquiv)
        dlg.type_keys(caminho)
        time.sleep(1)
        dlg.type_keys('{ENTER}')



