import json

class Config:
    def __init__(self):
        self.config = None
        self.popula_config()
        
    def get_config(self):
        return self.config
    
    def set_config(self, config):
        self.config = config
        with open("config.json", 'w', encoding="utf-8") as file:
            json.dump(config, file, indent=4, ensure_ascii=False)
    
    def popula_config(self):
        #no bloco abaixo try except ele tenta ler o arquivo config.json. se não houver cria um padrão.
        try:
            with open("config.json", encoding="utf-8") as file:
                self.config = json.load(file)
        except FileNotFoundError:
            self.default()
    
    #Cria o config.json padrão caso este não exista.
    def default(self):
        config_json = {
            "ultimo_acesso_user": "",
            "ultimo_orgao": "",
            "url_sei": "https://sei.economia.gov.br",
            "url_siapenet": "https://www1.siapenet.gov.br/orgao/Login.do?method=inicio",
            "unidade_sei": "MGI-SGP-DECIPEX-COATE-AFD",

            "lista_orgaos": [
                "MGI", "ME", "CMB", "COAF", "MTP", "MF", "MPO", "MDIC", "MPI", "MPS", "MEMP"
            ],

            "vinculos_decipex": [
                "40802", "40805", "40806"
            ]            
        }
        
        try:
            with open("config.json", 'w', encoding="utf-8") as file:
                json.dump(config_json, file, indent=4, ensure_ascii=False)
            self.config = config_json
        except Exception as e:
            print(f"Erro ao criar configuração padrão: {e}")
