import os
import glob
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


class SeleniumSetup:

    def __init__(self):
        self.base_path = os.path.dirname(os.path.abspath(__file__))

        self.download_folder = os.path.join(self.base_path, 'downloads')

        # Criar o diretório 'downloads' se não existir
        if not os.path.exists(self.download_folder):
            os.makedirs(self.download_folder)

        # Limpar a pasta de downloads
        self.limpar_downloads(self.download_folder)

    @property
    def get_download_folder(self):
        return self.download_folder

    def limpar_downloads(self, pasta):
        """Remove arquivos com padrão hodcivws*.jsp da pasta especificada"""
        for arquivo in glob.glob(os.path.join(pasta, 'hodcivws*.jsp')):
            os.remove(arquivo)

    def setup_driver(self):
        options = Options()
        prefs = {
            "download.default_directory": self.download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "credentials_enable_service": False,
            "password_manager_enabled": False
        }
        # Removido o caminho binário do Chrome, já que vamos usar o navegador instalado
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_argument('disable-infobars')

        # Usando ChromeDriverManager para gerenciar o driver
        navegador = webdriver.Chrome(options=options)

        navegador.maximize_window()
        return navegador
