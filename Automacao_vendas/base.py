import pandas as pd
import os

# CAMINHO DA PASTA

caminho = r"C:\Users\rodrigo.souza\nlfrutas.com.br\NL_File-Server - Documentos\Comercial_Terceirizadas\CONSOLIDA_VENDAS\Vendas"

# ARQUIVOS

braga = "Braga.xlsx"
cereais_silveira = "Cereais Silveira.xlsx"
cidade = "Supermercados Cidade.xlsx"
consul = "Consul.xlsx"
cordeiro = "Cordeiro.xlsx"
dupovo = "Rede Dupovo.xlsx"
escola = "Supermercado Escola.xlsx"
levate = "Levate.xlsx"
pagpouco = "Smart Pagpouco.xlsx"
pejoal = "Supervarejista Pejoal.xlsx"
rei_do_arroz = "Rei do Arroz.xlsx"
rena = "Rena.xlsx"
superminas = "Superminas.xlsx"

# LEITURA DOS ARQUIVOS

df_braga = pd.read_excel(os.path.join(caminho, braga), sheet_name="comparativo", header=2)
df_cereais_silveira = pd.read_excel(os.path.join(caminho, cereais_silveira), sheet_name="comparativo", header=2)
df_cidade = pd.read_excel(os.path.join(caminho, cidade), sheet_name="comparativo", header=2)
df_consul = pd.read_excel(os.path.join(caminho, consul), sheet_name="comparativo", header=2)
df_cordeiro = pd.read_excel(os.path.join(caminho, cordeiro), sheet_name="comparativo", header=2)
df_dupovo = pd.read_excel(os.path.join(caminho, dupovo), sheet_name="comparativo", header=2)
df_escola = pd.read_excel(os.path.join(caminho, escola), sheet_name="comparativo", header=2)
df_levate = pd.read_excel(os.path.join(caminho, levate), sheet_name="comparativo", header=2)
df_pagpouco = pd.read_excel(os.path.join(caminho, pagpouco), sheet_name="comparativo", header=2)
df_pejoal = pd.read_excel(os.path.join(caminho, pejoal), sheet_name="comparativo", header=2)
df_rei_do_arroz = pd.read_excel(os.path.join(caminho, rei_do_arroz), sheet_name="comparativo", header=2)
df_rena = pd.read_excel(os.path.join(caminho, rena), sheet_name="comparativo", header=2)
df_superminas = pd.read_excel(os.path.join(caminho, superminas), sheet_name="comparativo", header=2)


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

def get_diferenca(df, lista):

    if "Dif" not in df.columns:
        print("Coluna 'Dif' não encontrada no DataFrame. Verifique a origem dos dados.")
        return
    
    # Converte para numérico se existir
    df["Dif"] = pd.to_numeric(df["Dif"], errors="coerce")

    df_filtrado = df[((df["Dif"] > 0.009) | (df["Dif"] < -0.0099)) & (~df["Data"].isna())]
   
    return pd.concat([lista, df_filtrado], ignore_index=True)

diferencas = get_diferenca(df_braga, diferencas)
diferencas = get_diferenca(df_cereais_silveira, diferencas)
diferencas = get_diferenca(df_cidade, diferencas)
diferencas = get_diferenca(df_consul, diferencas)
diferencas = get_diferenca(df_cordeiro, diferencas)
diferencas = get_diferenca(df_dupovo, diferencas)
diferencas = get_diferenca(df_escola, diferencas)
diferencas = get_diferenca(df_levate, diferencas)
diferencas = get_diferenca(df_pagpouco, diferencas)
diferencas = get_diferenca(df_pejoal, diferencas)
diferencas = get_diferenca(df_rei_do_arroz, diferencas)
diferencas = get_diferenca(df_rena, diferencas)
diferencas = get_diferenca(df_superminas, diferencas)

diferencas = diferencas.dropna(axis=1, how='all')

try:

    df = diferencas.dropna(subset=["Parceiro", "Data", "Dif"])
    df["Data"] = pd.to_datetime(df["Data"]).dt.date

except:
    df = diferencas



