import re
from datetime import datetime

class Aposentados:
    def __init__(self, linha_planilha, status, status_cacoaposse, nome, cpf, siape, vinculo_decipex, orgao_origem, data_aposentadoria, data_dou, fundamento_legal, dl_aposentadoria):
        self.linha_planilha     = linha_planilha
        self.status             = status
        self.status_cacoaposse  = status_cacoaposse 
        self.nome               = nome
        self.cpf                = self.trata_cpf(cpf)
        self.siape              = siape
        self.vinculo_decipex    = vinculo_decipex        
        self.orgao_origem       = orgao_origem
        self.data_aposentadoria = self.set_data_aposentadoria(data_aposentadoria)
        self.fundamento_legal   = fundamento_legal
        self.dl_aposentadoria   = dl_aposentadoria
        self.data_dou           = self.set_data_dou(data_dou)

    def set_data_dou(self, data_dou):
        self.data_dou = self.trata_data(data_dou)

    def set_data_aposentadoria(self, data_aposentadoria):
        self.data_aposentadoria = self.trata_data(data_aposentadoria)


    def trata_cpf(self, cpf):
        '''Tratamento cpf garantindo que sempre tenha 11 dígitos'''
        # Remover caracteres não numéricos
        cpf = ''.join(filter(str.isdigit, str(cpf)))
        # garante que o cpf tenha 11 digitos adiciona zeros a esquerda
        return cpf.zfill(11)


    @staticmethod
    def trata_data(data):
        '''Converte uma data no formato ddMMMyyyy para dd/MM/yyyy'''

        # Define o padrão para ddMMMyyyy
        padrao_data = r'\d{2}[a-zA-Z]{3}\d{4}'

        # Tenta encontrar a data no texto
        resultado = re.findall(padrao_data, data)

        if resultado:
            data_encontrada = resultado[0]
            try:
                # Converte a data encontrada para o formato dd/MM/yyyy
                data_convertida = datetime.strptime(data_encontrada, '%d%b%Y')
                data_formatada = data_convertida.strftime('%d/%m/%Y')
                data_formatada = f'DATA APOSENTADORIA: {data_formatada}'
                return data_formatada
            except ValueError:
                return data
        else:
            return data
        
    def __str__(self):
        return (f"Aposentados(linha_planilha={self.linha_planilha}, status={self.status}, "
                f"status_cacoaposse={self.status_cacoaposse}, nome={self.nome}, cpf={self.cpf}, "
                f"siape={self.siape}, vinculo_decipex={self.vinculo_decipex}, orgao_origem={self.orgao_origem}, "
                f"data_aposentadoria={self.data_aposentadoria}, data_dou={self.data_dou}, "
                f"fundamento_legal={self.fundamento_legal}, dl_aposentadoria={self.dl_aposentadoria})")

    