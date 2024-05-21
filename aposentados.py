class Aposendados:
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



    def trata_cpf(cpf):
        '''tratamento cpf garantindo que sempre tenha 11 digitos'''        
        return cpf
    
    def trata_data(data):
        '''tratamento data para ficar no padrão dd/mm/aaaa'''
        return