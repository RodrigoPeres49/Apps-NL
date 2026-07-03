import pandas as pd
import os

caminho = r"C:\Users\rodrigo.souza\nlfrutas.com.br\NL_File-Server - Documentos\Comercial_Terceirizadas\CONSOLIDA_VENDAS\Vendas\A Controle de importações\Acerto de valor - Importação\Relatorio.xlsx"
caminho_arquivos = r"C:\Users\rodrigo.souza\nlfrutas.com.br\NL_File-Server - Documentos\Comercial_Terceirizadas\CONSOLIDA_VENDAS\Vendas\A Controle de importações\Acerto de valor - Importação"

# CRIANDO VARIÁVEIS DOS RELATÓRIOS

relatorios = pd.read_excel(caminho, sheet_name=None, dtype={"CNPJ": str, "CODPROD": str})
    
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
        relatorio["VLRVENDA"] = pd.to_numeric(relatorio["VLRVENDA"], errors="coerce")
        
        # REMOVE ZEROS
        relatorio = relatorio[relatorio["VLRVENDA"] != 0]
        relatorio = relatorio[relatorio["QTD"] != 0]
        
    
        relatorio["VLRVENDA"] = (
            relatorio["VLRVENDA"]
            .fillna(0)
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


    pasta_saida = os.path.join(caminho_arquivos, "Relatorios_Gerados", nome_loja)
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

print(relatorios)
os.system("pause")

for nome, df in relatorios.items():
    print(f"\nProcessando {nome}...")
    gerar_relatorio(df, nome)

os.system("pause")
