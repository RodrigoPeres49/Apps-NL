import pandas as pd
import os
import warnings
warnings.simplefilter(action='ignore', category=UserWarning)

# CAMINHO DA PASTA

caminho = r"C:\Users\rodrigo.souza\nlfrutas.com.br\NL_File-Server - Documentos\Comercial_Terceirizadas\CONSOLIDA_VENDAS\Vendas"

# ARQUIVOS

braga = "Braga.xlsx"
cereais_silveira = "Cereais Silveira.xlsx"
cidade = "Supermercados Cidade.xlsx"
consul = "Consul.xlsx"
cordeiro = "Cordeiro.xlsx"
donadio = "Donadio.xlsx"
dupovo = "Rede Dupovo.xlsx"
escola = "Supermercado Escola.xlsx"
levate = "Levate.xlsx"
pagpouco = "Smart Pagpouco.xlsx"
pejoal = "Supervarejista Pejoal.xlsx"
rei_do_arroz = "Rei do Arroz.xlsx"
rena = "Rena.xlsx"
superminas = "Superminas.xlsx"
sistema = "VENDAS_TERCEIRIZADAS.xls"

# LEITURA DOS ARQUIVOS

df_braga = pd.read_excel(os.path.join(caminho, braga), sheet_name="comparativo", header=2)
df_cereais_silveira = pd.read_excel(os.path.join(caminho, cereais_silveira), sheet_name="comparativo", header=2)
df_cidade = pd.read_excel(os.path.join(caminho, cidade), sheet_name="comparativo", header=2)
df_consul = pd.read_excel(os.path.join(caminho, consul), sheet_name="comparativo", header=2)
df_cordeiro = pd.read_excel(os.path.join(caminho, cordeiro), sheet_name="comparativo", header=2)
df_donadio = pd.read_excel(os.path.join(caminho, donadio), sheet_name="comparativo", header=2)
df_dupovo = pd.read_excel(os.path.join(caminho, dupovo), sheet_name="comparativo", header=2)
df_escola = pd.read_excel(os.path.join(caminho, escola), sheet_name="comparativo", header=2)
df_levate = pd.read_excel(os.path.join(caminho, levate), sheet_name="comparativo", header=2)
df_pagpouco = pd.read_excel(os.path.join(caminho, pagpouco), sheet_name="comparativo", header=2)
df_pejoal = pd.read_excel(os.path.join(caminho, pejoal), sheet_name="comparativo", header=2)
df_rei_do_arroz = pd.read_excel(os.path.join(caminho, rei_do_arroz), sheet_name="comparativo", header=2)
df_rena = pd.read_excel(os.path.join(caminho, rena), sheet_name="comparativo", header=2)
df_superminas = pd.read_excel(os.path.join(caminho, superminas), sheet_name="comparativo", header=2)
df_sistema = pd.read_excel(os.path.join(caminho,sistema), header=2)
df_sistema["Parceiro"] = df_sistema["Parceiro"].astype("Int64")
df_sistema["Data da Venda"] = pd.to_datetime(df_sistema["Data da Venda"], errors="coerce").dt.date

# PREPARANDO ARQUIVOS PARA AUTOMAÇÃO 

def definir_cabecalho(df, new_header = 0):
    new_header = df.iloc[0]
    df = df[1:]
    df.columns = new_header
    return df 


# DEFININDO CABEÇALHO EM TODOS OS ARQUIVOS

df_braga = definir_cabecalho(df_braga)
df_cereais_silveira = definir_cabecalho(df_cereais_silveira)
df_cidade = definir_cabecalho(df_cidade)
df_consul = definir_cabecalho(df_consul)
df_cordeiro = definir_cabecalho(df_cordeiro)
df_donadio = definir_cabecalho(df_donadio)
df_dupovo = definir_cabecalho(df_dupovo)
df_escola = definir_cabecalho(df_escola)
df_levate = definir_cabecalho(df_levate)
df_pagpouco = definir_cabecalho(df_pagpouco)
df_pejoal = definir_cabecalho(df_pejoal)
df_rei_do_arroz = definir_cabecalho(df_rei_do_arroz)
df_rena = definir_cabecalho(df_rena)
df_superminas = definir_cabecalho(df_superminas)

# VARIAVEL PARA ARMAZENAR LISTA DE DIFERENÇAS

diferencas = pd.DataFrame()

# DEFININDO LINHAS DE TODAS AS LOJAS QUE CONTEM DIFERENÇAS

def get_diferenca(df, sistema, lista):
    # Garantir tipos iguais
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.date
    sistema["Data da Venda"] = pd.to_datetime(sistema["Data da Venda"], errors="coerce").dt.date
    
    df["Parceiro"] = pd.to_numeric(df["Parceiro"], errors="coerce").astype("Int64")
    sistema["Parceiro"] = pd.to_numeric(sistema["Parceiro"], errors="coerce").astype("Int64")

    # Merge pelas duas chaves
    df_merge = df.merge(
        sistema[["Parceiro", "Data da Venda"]],
        how="left",
        left_on=["Parceiro", "Data"],
        right_on=["Parceiro", "Data da Venda"],
        indicator=True
    )
    
    # Pega apenas registros que NÃO estão no sistema
    df_filtrado = df_merge[df_merge["_merge"] == "left_only"]
    
    # Remove colunas auxiliares
    df_filtrado = df_filtrado.drop(columns=["Data da Venda", "_merge"])
    
    return pd.concat([lista, df_filtrado], ignore_index=True)

diferencas = get_diferenca(df_braga, df_sistema, diferencas)
diferencas = get_diferenca(df_cereais_silveira, df_sistema, diferencas)
diferencas = get_diferenca(df_cidade, df_sistema, diferencas)
diferencas = get_diferenca(df_consul, df_sistema, diferencas)
diferencas = get_diferenca(df_cordeiro, df_sistema, diferencas)
diferencas = get_diferenca(df_donadio, df_sistema, diferencas)
diferencas = get_diferenca(df_dupovo, df_sistema, diferencas)
diferencas = get_diferenca(df_escola, df_sistema, diferencas)
diferencas = get_diferenca(df_levate, df_sistema, diferencas)
diferencas = get_diferenca(df_pagpouco, df_sistema, diferencas)
diferencas = get_diferenca(df_pejoal, df_sistema, diferencas)
diferencas = get_diferenca(df_rei_do_arroz, df_sistema, diferencas)
diferencas = get_diferenca(df_rena, df_sistema, diferencas)
diferencas = get_diferenca(df_superminas, df_sistema, diferencas)

diferencas = diferencas.dropna(axis=1, how='all')

try:

    df = diferencas.dropna(subset=["Parceiro", "Data", "Dif"])
    df["Data"] = pd.to_datetime(df["Data"]).dt.date

except:
    df = diferencas







