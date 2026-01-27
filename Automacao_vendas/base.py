import pandas as pd
import os
from datetime import datetime
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

def planilha_aberta(e):
    os.system("cls")
    print("Houve um erro ao iniciar a aplicação:")
    print("\n - Verifique se existe planilhas abertas.\n - Salve as mesmas e feche todas as planilhas e incie novamente a aplicação.\n")
    print(f"Erro:\n\n {e}\n\n")
    os.system("pause")
    
    


# CAMINHO DA PASTA

caminho = r"C:\Users\rodrigo.souza\nlfrutas.com.br\NL_File-Server - Documentos\Comercial_Terceirizadas\CONSOLIDA_VENDAS\Vendas"

# ARQUIVOS

arquivos = {
    "Badiao": "Badiao.xlsx",
    "Braga": "Braga.xlsx",
    "Cereais Silveira": "Cereais Silveira.xlsx",
    "Supermercados Cidade": "Supermercados Cidade.xlsx",
    "Consul": "Consul.xlsx",
    "Cordeiro": "Cordeiro.xlsx",
    "Donadio": "Donadio.xlsx",
    "Rede Dupovo": "Rede Dupovo.xlsx",
    "Supermercado Escola": "Supermercado Escola.xlsx",
    "Levate": "Levate.xlsx",
    "Smart Pagpouco": "Smart Pagpouco.xlsx",
    "Supervarejista Pejoal": "Supervarejista Pejoal.xlsx",
    "Rei do Arroz": "Rei do Arroz.xlsx",
    "Rena": "Rena.xlsx",
    "Superminas": "Superminas.xlsx",
    "Supermercado Vidal": "Supermercado Vidal.xlsx",
}

# ARQUIVO DO CONTROLE DE LANÇAMENTOS E PENDÊNCIAS

controle_pendencias = "Vendas Terceirizadas Lançamentos.xlsx"

# ARQUIVO DAS VENDAS LANÇADAS NO SISTEMA

sistema = "VENDAS_TERCEIRIZADAS.xls"


try:

    df_sistema = pd.read_excel(os.path.join(caminho,sistema), header=2) 
    df_sistema["Parceiro"] = df_sistema["Parceiro"].astype("Int64") 
    df_sistema["Data da Venda"] = pd.to_datetime(df_sistema["Data da Venda"], errors="coerce").dt.date

    setores_remover = ["CONSUMO INTERNO", "SEM REPASSE", "VALOR TOTAL LOJA MENSAL", "DESCONTO MERCADORIA"]
    df_sistema = df_sistema[~df_sistema["Setor da Loja"].isin(setores_remover)]

except PermissionError as e:

    planilha_aberta()


# FUNCAO PARA PEGAR DATA DE MODIFICAÇÃO DO ARQUIVO

try:

    def get_datamodificacao(caminho,arquivo):
        caminho_completo = os.path.join(caminho,arquivo)
    
        if os.path.exists(caminho_completo):
            timestamp = os.path.getmtime(caminho_completo)
            data_modificacao = datetime.fromtimestamp(timestamp)
            return data_modificacao.strftime("%d/%m/%Y %H:%M:%S")
        else: 
            return "Arquivo não encontrado."


    # FUNCAO PARA ABRIR O ARQUIVO PARA UTILIZAR NA INTERFACE

    def abrir_arquivo(caminho, arquivo):
        try:
            os.startfile(os.path.join(caminho, arquivo))
            return True
        except Exception as e:
            print(f"Erro ao abrir o arquivo: {e}")
            return False
    
    # LEITURA DOS ARQUIVOS
    
    dfs = {}  # dicionário para armazenar DataFrames
    
    for nome, arquivo in arquivos.items():
        caminho_completo = os.path.join(caminho, arquivo)
        dfs[nome] = pd.read_excel(caminho_completo, sheet_name="comparativo", header=2)
    
except PermissionError as e:

    planilha_aberta(e)

# PREPARANDO ARQUIVOS PARA AUTOMAÇÃO 

def definir_cabecalho(df):
    new_header = df.iloc[0]
    df = df[1:].copy()  
    df.columns = new_header
    return df


# DEFININDO CABEÇALHO EM TODOS OS ARQUIVOS

df_badiao = definir_cabecalho(dfs["Badiao"])
df_braga = definir_cabecalho(dfs["Braga"])
df_cereais_silveira = definir_cabecalho(dfs["Cereais Silveira"])
df_cidade = definir_cabecalho(dfs["Supermercados Cidade"])
df_consul = definir_cabecalho(dfs["Consul"])
df_cordeiro = definir_cabecalho(dfs["Cordeiro"])
df_donadio = definir_cabecalho(dfs["Donadio"])
df_dupovo = definir_cabecalho(dfs["Rede Dupovo"])
df_escola = definir_cabecalho(dfs["Supermercado Escola"])
df_levate = definir_cabecalho(dfs["Levate"])
df_pagpouco = definir_cabecalho(dfs["Smart Pagpouco"])
df_pejoal = definir_cabecalho(dfs["Supervarejista Pejoal"])
df_rei_do_arroz = definir_cabecalho(dfs["Rei do Arroz"])
df_rena = definir_cabecalho(dfs["Rena"])
df_superminas = definir_cabecalho(dfs["Superminas"])
df_vidal = definir_cabecalho(dfs["Supermercado Vidal"])


# VARIAVEL PARA ARMAZENAR LISTA DE DIFERENÇAS

diferencas = pd.DataFrame()

# DEFININDO LINHAS DE TODAS AS LOJAS QUE CONTEM DIFERENÇAS

def get_diferenca(df, sistema, lista):

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    sistema["Data da Venda"] = pd.to_datetime(sistema["Data da Venda"], errors="coerce")
    
    df["Parceiro"] = pd.to_numeric(df["Parceiro"], errors="coerce").astype("Int64")
    sistema["Parceiro"] = pd.to_numeric(sistema["Parceiro"], errors="coerce").astype("Int64")

    df_merge = df.merge(
        sistema[["Parceiro", "Data da Venda"]],
        how="left",
        left_on=["Parceiro", "Data"],
        right_on=["Parceiro", "Data da Venda"],
        indicator=True
    )
    

    df_filtrado = df_merge[df_merge["_merge"] == "left_only"]
    

    df_filtrado = df_filtrado.drop(columns=["Data da Venda", "_merge"])
    
    return pd.concat([lista, df_filtrado], ignore_index=True)

diferencas = get_diferenca(df_badiao, df_sistema, diferencas)
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
diferencas = get_diferenca(df_vidal, df_sistema, diferencas)

diferencas = diferencas.dropna(axis=1, how='all')

try:

    df = diferencas.dropna(subset=["Parceiro", "Data", "Dif"])
    df["Data"] = pd.to_datetime(df["Data"]).dt.date

except:
    df = diferencas











