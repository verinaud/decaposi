from datetime import datetime

class Aposentados:
    def __init__(self, linha_planilha, nome, cpf, vinculo_decipex, siape, orgao_origem, data_aposentadoria, fundamento_legal, num_portaria, data_dou):
        self.linha_planilha     = linha_planilha
        self.nome               = nome
        self.cpf                = self.trata_cpf(cpf)
        self.vinculo_decipex    = vinculo_decipex
        self.siape              = siape
        self.orgao_origem       = orgao_origem
        self.data_aposentadoria = self.trata_data(data_aposentadoria)
        self.fundamento_legal   = fundamento_legal
        self.num_portaria       = num_portaria
        self.data_dou           = self.trata_data(data_dou)

    def trata_cpf(self, cpf):
        '''Tratamento cpf garantindo que sempre tenha 11 dígitos'''
        # Remover caracteres não numéricos
        cpf = ''.join(filter(str.isdigit, str(cpf)))
        # garante que o cpf tenha 11 digitos
        return cpf.zfill(11)

    def trata_data(self, data):
        '''Tratamento de data para ficar no padrão dd/mm/aaaa'''
        if isinstance(data, str):
            # Tentar converter a string para um objeto datetime
            try:
                data = datetime.strptime(data, '%d/%m/%Y')
            except ValueError:
                # Se a string não estiver no formato esperado, retornar como está
                return data
        if isinstance(data, datetime):
            # Se for um objeto datetime, formatar como string
            return data.strftime('%d/%m/%Y')
        return data