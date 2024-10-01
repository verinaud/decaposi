import re
from datetime import datetime

class Aposentado:
    def __init__(self, linha_planilha, status, cpf, nome, siape, vinculo_decipex, orgao_origem, data_aposentadoria, data_dou, fundamento_legal, dl_aposentadoria):
        self.linha_planilha     = linha_planilha
        self.status             = status
        self.cpf                = self.trata_cpf(cpf)
        self.nome               = nome
        self.siape              = siape
        self.vinculo_decipex    = vinculo_decipex        
        self.orgao_origem       = orgao_origem
        self.data_aposentadoria = data_aposentadoria
        self.data_dou           = data_dou
        self.fundamento_legal   = fundamento_legal
        self.dl_aposentadoria   = dl_aposentadoria
        

    def set_data_dou(self, data_dou):
        self.data_dou = self.trata_data(data_dou)

    def set_data_aposentadoria(self, data_aposentadoria):
        self.data_aposentadoria = self.trata_data(data_aposentadoria)


    def trata_cpf(self, cpf):
        from tela_mensagem import Mensagem as msg
        cpf = str(cpf)
        # Remove pontos e traços
        cpf_limpo = cpf.replace(".", "").replace("-", "")

        # Garante que tenha 11 dígitos sem alterar se já tiver 11
        if len(cpf_limpo) < 11:
            cpf_formatado = cpf_limpo.zfill(11)
        else:
            cpf_formatado = cpf_limpo
        return cpf_formatado


    def trata_data(self, data):
        if data is not None:
            try:
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
                        return data_formatada
                    except ValueError:
                        return data_formatada
                else:
                    return data
            except Exception as erro:
                #self.status  = "Verificação Manual"
                return "Verificação Manual"
        else:
            return None
        
    def __str__(self):
        return (f"Aposentados(linha_planilha={self.linha_planilha}, status={self.status}, "
                f"siape={self.siape}, vinculo_decipex={self.vinculo_decipex}, orgao_origem={self.orgao_origem}, "
                f"data_aposentadoria={self.data_aposentadoria}, data_dou={self.data_dou}, "
                f"fundamento_legal={self.fundamento_legal}, dl_aposentadoria={self.dl_aposentadoria})")

    