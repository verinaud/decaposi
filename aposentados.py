import re
from datetime import datetime

class Aposentados:
    def __init__(self, linha_planilha, status, status_cacoaposse, nome, cpf, siape, vinculo_decipex, orgao_origem, data_aposentadoria, data_dou, fundamento_legal, dl_aposentadoria, data_inicio):
        self.linha_planilha     = linha_planilha
        self.status             = status
        self.status_cacoaposse  = status_cacoaposse 
        self.nome               = nome
        self.cpf                = self.trata_cpf(cpf)
        self.siape              = siape
        self.vinculo_decipex    = vinculo_decipex        
        self.orgao_origem       = orgao_origem
        self.data_aposentadoria = self.trata_data(data_aposentadoria)
        self.fundamento_legal   = fundamento_legal
        self.dl_aposentadoria   = dl_aposentadoria
        self.data_inicio        = self.trata_data(data_inicio)
        self.data_dou           = self.extrair_data(data_dou)

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
                data_formatada = f'DATA INICIO: {data_formatada}'
                return data_formatada
            except ValueError:
                return data
        else:
            return data

    @staticmethod
    def extrair_data(texto):
        '''Extrai a última data do texto fornecido e retorna no formato 'DATA DOU: DD/MM/YYYY' '''
        padrao_data1 = r'\d{2}/\d{2}/\d{4}'
        padrao_data2 = r'\d{2}[a-zA-Z]{3}\d{4}'
        
        # Encontrar todas as datas no texto
        resultados1 = re.findall(padrao_data1, texto)
        resultados2 = re.findall(padrao_data2, texto)
        
        # Combinar todas as datas encontradas
        todas_as_datas = resultados1 + resultados2
        
        if not todas_as_datas:
            return 'DATA DOU: Desconhecida'
        
        # Ordenar todas as datas para garantir que a mais recente seja a última
        todas_as_datas.sort(reverse=True)
        
        data_formatada = todas_as_datas[0]
        
        # Converter o formato de data
        if len(data_formatada) == 10:  # Formato DD/MM/YYYY
            data_formatada = f'DATA DOU: {data_formatada}'
        elif len(data_formatada) == 9:  # Formato DDMMMYYYY
            try:
                data_formatada = datetime.strptime(data_formatada, '%d%b%Y').strftime('DATA DOU: %d/%m/%Y')
            except ValueError:
                data_formatada = 'DATA DOU: Desconhecida'
        
        return data_formatada