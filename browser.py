from webdriver_manager.chrome           import ChromeDriverManager
from selenium.webdriver.chrome.options  import Options
from selenium.webdriver.chrome.service  import Service
from selenium                           import webdriver
import os

class Browser:
    def __init__(self):
        self.__browser = None        
    
    def get_browser_01(self):
        # Instancia a classe Options
        chrome_options = Options()

        chrome_options.add_experimental_option(
            'prefs', {
                "profile.password_manager_enabled": False,
                "credentials_enable_service": False,
                "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":"","name":"Save as PDF"}],"selectedDestinationId":"Save as PDF","version":2}',
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True
            }
        )

        # Oculta a informação de que o navegador está sendo controlado por automação
        chrome_options.add_argument('disable-infobars')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-password-manager-reauthentication")
        chrome_options.add_argument("--disable-save-password-bubble")

        # Inicializa o browser com o Chrome Driver e passa as opções definidas por parâmetro
        service = Service(ChromeDriverManager().install())
        self.__browser = webdriver.Chrome(options=chrome_options, service=service)

        

        return self.__browser

    def get_browser_02(self):
        # Instancia a classe Options
        chrome_options = Options()

        # Define as opções para o navegador (browser)
        chrome_options.add_experimental_option(
            'prefs', {
                "profile.password_manager_enabled": False,
                "credentials_enable_service": False,
                "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":"","name":"Save as PDF"}],"selectedDestinationId":"Save as PDF","version":2}',
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True
            }
        )

        # Modo quiosque de impressão
        chrome_options.add_argument('--kiosk-printing') 

        # Oculta a informação de que o navegador está sendo controlado por automação
        chrome_options.add_argument('disable-infobars')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-password-manager-reauthentication")
        chrome_options.add_argument("--disable-save-password-bubble")

        # Inicializa o browser com o Chrome Driver e passa as opções definidas por parâmetro
        service = Service(ChromeDriverManager().install())
        self.__browser = webdriver.Chrome(options=chrome_options, service=service)

        return self.__browser    
    
    def get_browser_03(self, diretorio_download):
        # Define o diretório de download
        download_path       = diretorio_download

        # Cria o diretório se não existir
        if not os.path.exists(download_path):
            os.makedirs(download_path)

        # Instancia a classe Options
        chrome_options = Options()

        # Define as opções para o navegador (browser)
        chrome_options.add_experimental_option(
            'prefs', {
                "profile.password_manager_enabled": False,
                "credentials_enable_service": False,
                "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":"","name":"Save as PDF"}],"selectedDestinationId":"Save as PDF","version":2}',
                "savefile.default_directory": download_path,
                "download.default_directory": download_path,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True
            }
        )

        # Modo quiosque de impressão
        chrome_options.add_argument('--kiosk-printing') 

        # Oculta a informação de que o navegador está sendo controlado por automação
        chrome_options.add_argument('disable-infobars')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])

        # Inicializa o browser com o Chrome Driver e passa as opções definidas por parâmetro
        service = Service(ChromeDriverManager().install())
        self.__browser = webdriver.Chrome(options=chrome_options, service=service)

        return self.__browser