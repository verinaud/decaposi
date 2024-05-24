import sys
from configuracao import Config
from interface import Interface
import pandas as pd
from aposentados import Aposentados

class Decaposi():
    def __init__(self):
        self.config = Config() # Cria uma instancia da classe Config

        self.json = self.config.get_json() # Pega o config_json

        self.interface = Interface() # Cria uma instancia da classe Interface

        self.window = self.interface.window # Recebe objeto UI da classe Interface

        self.lista_aposentados = []

        self.evento_botoes()        

        sys.exit(self.interface.app.exec_()) # inicia o loop de eventos da aplicação PyQt    

    def iniciar(self):
        print("Iniciou")                       

        self.lista_aposentados = self.ler_base_dados() #Beatriz
        for ls in self.lista_aposentados:
            print(ls.nome, ls.cpf, ls.siape)

        self.consultar_vinculo_decipex() #Tem pronto

        self.consultar_cacoaposse() #Beatriz

        self.preencher_declaracao() #Beatriz/André
    
    
    def ler_base_dados(self, nome=None, cpf=None, siape=None):
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
            1. Procura pela planilha base_dados_solicitacao.xlsx na raiz do programa,
               se não encontrar solicita que o usuário selecione no computador.
            2. Lê os dados da planilha e instancia a classe Aposentados
            3. Cria uma lista com objetos da classe Aposentados
            4. Renomeia a planilha como "base_dados_solicitacao.xlsx"
            5. Retorna uma lista de objetos da classe Aposentados
        """
        lista = []

        try:
            # Ler o arquivo Excel e armazenar o resultado em um DataFrame
            df = pd.read_excel('Planilha de Declarações Aposentados.xlsx')

            # Aplicar filtros se os parâmetros foram fornecidos
            if nome:
                df = df[df['Nome'] == nome]
            if cpf:
                df = df[df['CPF'] == cpf]
            if siape:
                df = df[df['SIAPE'] == siape]

            # Iterar sobre as linhas do DataFrame filtrado
            for _, row in df.iterrows():
                aposentado = Aposentados(
                    linha_planilha=row.name,
                    nome=row['Nome'],
                    cpf=row['CPF'],
                    vinculo_decipex=False,  # Valor padrão, será atualizado posteriormente
                    siape=row['SIAPE'],
                    orgao_origem=None,  # Valor padrão, será atualizado posteriormente
                    data_aposentadoria=None,  # Valor padrão, será atualizado posteriormente
                    fundamento_legal=None,  # Valor padrão, será atualizado posteriormente
                    num_portaria=None,  # Valor padrão, será atualizado posteriormente
                    data_dou=None  # Valor padrão, será atualizado posteriormente
                )
                lista.append(aposentado)

        except FileNotFoundError:
            print("Arquivo 'Planilha de Declarações Aposentados.xlsx' não encontrado. Por favor, verifique o nome do arquivo e sua localização.")

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
            3. Atualiza os seguintes parâmetros da lista_aposentados: vinculo_decipex com True ou False
            4. Fecha o terminal 3270
            5. atualiza a planilha excel acrescentando vinculo_decipex
        """
        

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

    

if __name__ == "__main__":
        bot = Decaposi()
        sys.exit()