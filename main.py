from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.chart import LineChart,Reference
from datetime import date
from func import *
#Pergunta qual açao o usuario quer procurar
#acao = input("Qual o código da ação? ").upper()
acao = "BIDI4"
#Abrindo o arquivo com as açoes e pegando a informoçao da açao necessária
leitor_acoes = LeitorAcoes(caminho_arquivo="./dados/")
leitor_acoes.processa_arquivo(acao)
#Invoca a classe GerenciadorPlanilha para colocarr os dados nas coolunas designadas
gerenciador = GerrenciadorPlanilha()
gerenciador.add_planilha("Dados")
gerenciador.add_linha(["Data","Cotação","Banda Inferior","Banda Superior"])

indice = 2
for linha in leitor_acoes.dados:
#Editando a coluna Data (2017-04-03 21:00:00;8.2606)
    ano_mes_dia = linha[0].split(" ")[0]
    data = date(
        year = int(ano_mes_dia.split("-")[0]),
        month = int(ano_mes_dia.split("-")[1]),
        day = int(ano_mes_dia.split("-")[2])
    )
#Cotacao
    cotacao = float(linha[1])
#Formulas das bandas de bollinger
    formula_bb_inferior = f"=AVERRAGE(B{indice}:B{indice + 19}) - 2*STDEV(B{indice}:B{indice + 19})"
    formula_bb_superior = f"=AVERRAGE(B{indice}:B{indice + 19}) - 2*STDEV(B{indice}:B{indice + 19})"
#Traz a função atualiza_celula para atualizar automaticamente cada celula
    gerenciador.atualiza_celula(celula=f"A{indice}", dado=data)
    gerenciador.atualiza_celula(celula=f"B{indice}", dado=cotacao)
    gerenciador.atualiza_celula(celula=f"C{indice}", dado=formula_bb_inferior)
    gerenciador.atualiza_celula(celula=f"D{indice}", dado=formula_bb_superior)
    indice += 1

gerenciador.add_planilha(titulo_planilha="Gráfico")

#Mesclagem de celulas para criação do cabeçalho do gráfico

cabecalho = planilha_grafico("A1")
cabecalho.font = Font(b=True, sz=18,color="FFFFFF")
cabecalho.fill = PatternFill("solid", fgColor="07838f")
cabecalho.aligment = Alignment(vertical="center", horizontal="center")
#Criando gráfico da planilha
grafico = LineChart()
grafico.width = 33.87
grafico.height = 14.82
grafico.title = f"Cotação - {acao}"
grafico.x_asis.title = "Data da Cotação"
grafico.y_assis.title = "Valor da Cotação"

referencia_cotacoes = Reference(planilha_ativa, min_col=2,min_row=2, max_col=4,max_row=indice)
referencia_datas = Reference(planilha_ativa, min_col=1,min_row=1, max_col=1,max_row=indice)
grafico.add_data(referencia_cotacoes)
grafico.set_categories(referencia_datas)

linha_cotacoes = grafico.series[0]
linha_bb_inferior = grafico.series[1]
linha_bb_superior = grafico.series[2]

linha_cotacoes.graphicalProperties.line.solidFill = "0a55ab"
linha_bb_inferior.graphicalProperties.line.width = 0
linha_bb_inferior.graphicalProperties.line.solidFill = "a61508"
linha_bb_superior.graphicalProperties.line.width = 0
linha_bb_superior.graphicalProperties.line.solidFill = "12a154"
planilha_grafico.add_chart(grafico,"A3")

workbook.save("./saida/Planilha.xlsx")
