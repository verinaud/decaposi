import estilos as estilo
from PyQt5.QtWidgets import QApplication, QMessageBox, QListWidget, QListWidgetItem, QComboBox, QHBoxLayout, \
    QHBoxLayout, QLineEdit, QDesktopWidget, QWidget, QVBoxLayout, QLabel, QPushButton, QMainWindow, QAction
from PyQt5.QtCore import QThread
from PyQt5.QtGui import QFont
import sys
from configuracao import Configuracao


class Interface():
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.app.setStyleSheet(estilo.dark_stylesheet)  # Aplica o estilo
        self.window = UI()  # Instancia a interface do usuário
        self.window.show()     
    

class UI(QMainWindow) :  # inserir QMainWindow no parâmetro
    
    def __init__(self) :
        super(UI, self).__init__()
        self.initUI()
                

    def initUI(self) :
        # Criar um único painel para ambos os widgets
        self.panel_container = QWidget(self)

        # Criar 8 widgets (painéis)
        self.panel1 = QWidget(self.panel_container)        

        # Adicionar widgets ao panel1
        # Campo de Login
        self.login_label_m = QLabel("Usuário e senha SEI.", self.panel1)
        
        self.login_label = QLabel("Usuário:", self.panel1)
        self.login_input = QLineEdit("", self.panel1)

        # Campo de senha
        self.password_label = QLabel("Senha:", self.panel1)
        self.password_input = QLineEdit(self.panel1)
        self.password_input.setEchoMode(QLineEdit.Password)

        self.show_password_button = QPushButton("Exibir", self.panel1)
        self.show_password_button.setCheckable(True)
        self.show_password_button.clicked.connect(self.toggle_password_visibility)
        self.show_password_button.setFixedWidth(100)

        # Campo de Órgão
        self.unidade_label = QLabel("Unidade:", self.panel1)

        # Criar um QComboBox
        self.unidade_combo_box = QComboBox(self.panel1)

        # Preenche o QComboBox com a lista de órgãos
        self.preencher_unidade_combo_box()
        

        #botão iniciar
        self.button_iniciar = QPushButton('Iniciar', self.panel1)
        #self.button_iniciar.clicked.connect(self.iniciar_coate)
        

        # Criar layouts verticais para os painéis
        layout1 = QVBoxLayout(self.panel1)
        

        # Adicionar os widgets aos layouts dos painéis
        layout1.addWidget(self.login_label_m)
        layout1.addWidget(self.login_label)
        layout1.addWidget(self.login_input)
        layout1.addSpacing(10)
        layout1.addWidget(self.password_label)
        layoutSenha = QHBoxLayout(self.panel1)
        layoutSenha.addWidget(self.password_input)
        layoutSenha.addWidget(self.show_password_button)
        layout1.addLayout(layoutSenha)
        layout1.addWidget(self.unidade_label)
        layout1.addWidget(self.unidade_combo_box)
        layout1.addSpacing(70)
        layout1.addWidget(self.button_iniciar)      
        

        # Criar um layout vertical
        layout = QVBoxLayout(self.panel_container)

        # Adicionar os painéis ao layout vertical
        layout.addWidget(self.panel1)
        
        # Esconder inicialmente o Painel 2
        self.panel1.show()
        

        # Adicionar o layout à janela principal
        self.panel_container.setLayout(layout)
        self.setCentralWidget(self.panel_container)

        menubar = self.menuBar()

        # Menu "Arquivo"
        exibir = menubar.addMenu('Exibir')

        # Item de menu "Janela principal"
        menu_jan_principal = QAction('Painel Home', self)
        menu_jan_principal.triggered.connect(self.showPanel1)
        exibir.addAction(menu_jan_principal)
        
        
        
        # Menu "Arquivo"
        ajuda = menubar.addMenu('Ajuda')       
        
        # Item de menu "Sobre"
        sobre_action = QAction('Sobre', self)
        sobre_action.triggered.connect(self.show_description)
        ajuda.addAction(sobre_action)
        
        

        # Exibir a janela
        altura = 400
        self.setGeometry(10, 40, 500, altura)

        self.panel_container.setMaximumHeight(altura-30)

        self.setWindowTitle('Declaração de aposentadoria')

        screen = QDesktopWidget().screenGeometry()

        window = self.geometry()

        self.move(int((screen.width() - window.width()) / 2), int((screen.height() - window.height()) / 2))

        self.show()

    def preencher_unidade_combo_box(self):
        config = Configuracao().get_config()
        self.unidade_combo_box.addItem(config["ultimo_orgao"])
        lista_orgaos = config.get("lista_orgaos", [])
        self.unidade_combo_box.addItems(lista_orgaos)


    def toggle_password_visibility(self):
        if self.show_password_button.isChecked():
            self.show_password_button.setText("Ocultar")
            self.password_input.setEchoMode(QLineEdit.Normal)
        else:
            self.show_password_button.setText("Exibir")
            self.password_input.setEchoMode(QLineEdit.Password)


    def showPanel1(self):
        self.panel1.show()
        #self.panel2.hide()
        #self.panel3.hide()
              
        self.setWindowTitle('Painel 1')

    def showPanel2(self):
        self.panel1.hide()
        #self.panel2.show()
        #self.panel3.hide()
              
        self.setWindowTitle('Painel 2')

    def showPanel3(self):
        self.panel1.hide()
        #self.panel2.hide()
        #self.panel3.show()
              
        self.setWindowTitle('Painel 3')

    def show_description(self):
        desc_text = """
        Sobre esta Automação:

        A presente automação tem por objetivo...

        """
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle("Descrição")
        msgBox.setText(desc_text)
        msgBox.exec_()