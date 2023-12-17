class LeitorAcoes:
    def __init__(self, caminho_arquivo:str = ""):
        self.caminho_arquivo = caminho_arquivo
        self.dados = []

    def processa_arquivo(self, acao:str):
        """
        
        """
        with open(f"{self.caminho_arquivo}{acao}.txt","r") as arquivo_cotacao:
            linhas = arquivo_cotacao.readlines()
            self.dados = [linha.replace("\n","").split(";") for linha in linhas]