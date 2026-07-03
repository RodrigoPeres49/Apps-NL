from base import *
import pandas as pd
import os

# CAMINHOS DOS ARQUIVOS  (OBS: O CAMINHO DA PASTA ESTÁ EM base.py)

caminho_badiao_acerto = os.path.join(caminho,arquivos["Badiao Acerto"])
caminho_badiao = os.path.join(caminho,arquivos["Badiao"])
caminho_braga = os.path.join(caminho, arquivos["Braga"])
caminho_cereais = os.path.join(caminho, arquivos["Cereais Silveira"])
caminho_cidade = os.path.join(caminho, arquivos["Cidade"])
caminho_consul = os.path.join(caminho, arquivos["Consul"])
caminho_cordeiro = os.path.join(caminho, arquivos["Cordeiro"])
caminho_donadio = os.path.join(caminho, arquivos["Donadio"])
caminho_dupovo = os.path.join(caminho, arquivos["Dupovo"])
caminho_levate = os.path.join(caminho, arquivos["Levate"])
caminho_pagpouco = os.path.join(caminho, arquivos["Pagpouco"])
caminho_pejoal = os.path.join(caminho, arquivos["Pejoal"])
caminho_reidoarroz = os.path.join(caminho, arquivos["Rei do Arroz"])
caminho_escola = os.path.join(caminho, arquivos["Escola"])
caminho_superminas = os.path.join(caminho, arquivos["Superminas"])
caminho_vidal = os.path.join(caminho, arquivos["Vidal"])

# caminho_rena = os.path.join(caminho, arquivos["Rena"])

# CRIANDO VARIÁVEIS DOS RELATÓRIOS

relatorios = {
    "Badiao Acerto": pd.read_excel(caminho_badiao_acerto, sheet_name="Formato Arquivo Badiao", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Uniao Acerto": pd.read_excel(caminho_badiao_acerto, sheet_name="Formato Arquivo Uniao", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Badiao": pd.read_excel(caminho_badiao, sheet_name="Formato Arquivo Badiao", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Uniao": pd.read_excel(caminho_badiao, sheet_name="Formato Arquivo Uniao", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Braga": pd.read_excel(caminho_braga, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Cereais Silveira": pd.read_excel(caminho_cereais, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Cidade": pd.read_excel(caminho_cidade, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Consul": pd.read_excel(caminho_consul, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Cordeiro": pd.read_excel(caminho_cordeiro, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Cordeiro Semanal": pd.read_excel(caminho_cordeiro, sheet_name="Formato Arquivo Semanal", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Donadio": pd.read_excel(caminho_donadio, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Dupovo": pd.read_excel(caminho_dupovo, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Levate": pd.read_excel(caminho_levate, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Pagpouco": pd.read_excel(caminho_pagpouco, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Pejoal": pd.read_excel(caminho_pejoal, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Rei do Arroz": pd.read_excel(caminho_reidoarroz, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Escola": pd.read_excel(caminho_escola, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Superminas": pd.read_excel(caminho_superminas, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
    "Vidal": pd.read_excel(caminho_vidal, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),

    # "Rena": pd.read_excel(caminho_rena, sheet_name="Formato Arquivo", header=3, dtype={"CNPJ": str, "CODPROD": str}),
}

nova_ordem = ["CNPJ","DATA","CODPROD","DESCRICAO","QTD", "VLRUNIT", "VLRVENDA"]

for nome, df in relatorios.items():
    relatorios[nome] = df.reindex(columns=nova_ordem)


# GERANDO RELATORIO

def gerar_relatorio(relatorio: pd.DataFrame, nome_loja: str):
    """Gera relatórios separados por CNPJ e Data de um DataFrame."""

    relatorio.columns = relatorio.columns.str.strip()


    colunas_necessarias = ["CNPJ", "DATA"]
    for col in colunas_necessarias:
        if col not in relatorio.columns:
            print(f"⚠️ Coluna '{col}' não encontrada em {nome_loja}.")
            return
    
    # REMOVENDO LINHAS COM VALOR TOTAL IGUAL A 0

    if "VLRVENDA" in relatorio.columns:

        relatorio["VLRVENDA"] = pd.to_numeric(
            relatorio["VLRVENDA"], errors="coerce"
        ).fillna(0)
    
        relatorio = relatorio[relatorio["VLRVENDA"] != 0]


        relatorio["VLRVENDA"] = (
        relatorio["VLRVENDA"]
        .astype(str)
        .str.replace(".", ",")
    )

        
    # CONVERTENDO VALORES E SUBSTITUINDO . POR ,

    for coluna in ["QTD", "VLRUNIT"]:
        if coluna in relatorio.columns:
            relatorio[coluna] = (
                pd.to_numeric(relatorio[coluna], errors="coerce") 
                .fillna(0)
                .astype(str)  
                .str.replace(".", ",")  
            )



    relatorio["DATA"] = pd.to_datetime(relatorio["DATA"], errors="coerce")

    # CORREÇÃO DE CNPJ E COD PRODUTO

    if "CNPJ" in relatorio.columns:
        def ajusta_cnpj(valor):
            
            valor = valor.replace(" ", "")

            if pd.isna(valor):
                return valor
            s = str(valor)

            if len(s) == 14:
                return s
            return "0" + s
    
    relatorio["CNPJ"] = relatorio["CNPJ"].apply(ajusta_cnpj)
    
    if "CODPROD" in relatorio.columns:
        relatorio["CODPROD"] = (
            relatorio["CODPROD"]
            .astype(str)
            .str.replace(".0", "", regex=False)
        )

    # PASTA AONDE SERÃO INSERIDOS OS RELATORIOS


    pasta_saida = os.path.join(caminho, "Relatorios_Gerados", nome_loja)
    os.makedirs(pasta_saida, exist_ok=True)
    
    # GARANTE QUE A DATA É DATETIME
    relatorio["DATA"] = pd.to_datetime(
        relatorio["DATA"], errors="coerce"
    )
    
    
    # GERANDO ARQUIVOS POR DATA
    for data, grupo in relatorio.groupby("DATA"):
        if pd.isna(data):
            continue
    
        # CRIA PASTA DO MÊS (YYYY-MM)
        pasta_mes = os.path.join(pasta_saida, data.strftime("%Y-%m"))
        os.makedirs(pasta_mes, exist_ok=True)
    
        nome_arquivo = f"{nome_loja}_{data.strftime('%Y-%m-%d')}.csv"
        caminho_arquivo = os.path.join(pasta_mes, nome_arquivo)
    
        # CONDIÇÃO PARA NÃO GERAR NOVAMENTE
        if os.path.exists(caminho_arquivo):
            print(f" Arquivo {nome_arquivo} já criado.")
        else:
            grupo.to_csv(caminho_arquivo, sep=";", index=False)
            print(f" Gerado: {caminho_arquivo}")
    

pd.options.mode.chained_assignment = None

for nome, df in relatorios.items():
    print(f"\nProcessando {nome}...")
    gerar_relatorio(df, nome)

os.system("pause")
