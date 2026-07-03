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

# OUTRAS FUNÇÕES

# ARQUIVO DO CONTROLE DE LANÇAMENTOS E PENDÊNCIAS

controle_pendencias = "Vendas Terceirizadas Lançamentos.xlsx"

# ARQUIVO PARA DO MANUAL DE UTILIZAÇÃO DOS LANÇAMENTOS

caminho_manual = r"C:\Users\rodrigo.souza\nlfrutas.com.br\NL_File-Server - Documentos\Comercial_Terceirizadas\CONSOLIDA_VENDAS\Vendas\Automacao_vendas\manual"
manual = "manual_lancamentos.docx"

# FUNCAO PARA ABRIR O ARQUIVO PARA UTILIZAR NA INTERFACE

def abrir_arquivo(caminho, arquivo):
    try:
        os.startfile(os.path.join(caminho, arquivo))
        return True
    except Exception as e:
        print(f"Erro ao abrir o arquivo: {e}")
        return False


# CAMINHO DA PASTA
caminho = r"C:\Users\rodrigo.souza\nlfrutas.com.br\NL_File-Server - Documentos\Comercial_Terceirizadas\CONSOLIDA_VENDAS\Vendas"

# ARQUIVOS
arquivos = {
    "Badiao Acerto": "Badiao Acerto.xlsx",
    "Badiao": "Badiao.xlsx",
    "Braga": "Braga.xlsx",
    "Cereais Silveira": "Cereais Silveira.xlsx",
    "Cidade": "Cidade.xlsx",
    "Consul": "Consul.xlsx",
    "Cordeiro": "Cordeiro.xlsx",
    "Donadio": "Donadio.xlsx",
    "Dupovo": "Dupovo.xlsx",
    "Escola": "Escola.xlsx",
    "Levate": "Levate.xlsx",
    "Pagpouco": "Pagpouco.xlsx",
    "Pejoal": "Pejoal.xlsx",
    "Rei do Arroz": "Rei do Arroz.xlsx",
    "Superminas": "Superminas.xlsx",
    "Vidal": "Supermercado Vidal.xlsx",
}

sistema = "VENDAS_TERCEIRIZADAS.xlsx"


# ===============================
# LEITURA DO SISTEMA
# ===============================

try:

    df_sistema = pd.read_excel(os.path.join(caminho, sistema), header=2)

    df_sistema["Parceiro"] = pd.to_numeric(df_sistema["Parceiro"], errors="coerce").astype("Int64")
    df_sistema["Data da Venda"] = pd.to_datetime(df_sistema["Data da Venda"], errors="coerce").dt.date

    setores_remover = [
        "CONSUMO INTERNO",
        "SEM REPASSE",
        "VALOR TOTAL LOJA MENSAL",
        "DESCONTO MERCADORIA"
    ]

    df_sistema = df_sistema[~df_sistema["Setor da Loja"].isin(setores_remover)]

    # AGRUPA O SISTEMA POR DIA
    df_sistema = (
        df_sistema
        .groupby(["Parceiro", "Data da Venda"], as_index=False)
        .agg({
            "Valor da Venda": "sum",
            "Quantidade Vendida": "sum"
        })
    )

except PermissionError as e:
    planilha_aberta(e)


# ===============================
# LEITURA DAS PLANILHAS
# ===============================

dfs = {}

for nome, arquivo in arquivos.items():
    caminho_completo = os.path.join(caminho, arquivo)
    dfs[nome] = pd.read_excel(caminho_completo, sheet_name="comparativo", header=2)


# ===============================
# LIMPEZA DAS PLANILHAS
# ===============================

def definir_cabecalho(df):

    new_header = df.iloc[0]
    df = df[1:].copy()
    df.columns = new_header

    # remove colunas vazias
    df = df.loc[:, df.columns.notna()]

    # remove linhas vazias
    df = df.dropna(how="all")

    # remove linhas de total
    df = df[df["Parceiro"].notna()]

    return df


# aplica em todas
for nome in dfs:
    dfs[nome] = definir_cabecalho(dfs[nome])


# ===============================
# FUNÇÃO DE DIFERENÇA
# ===============================

def get_diferenca(df, sistema, lista):

    df = df.copy()

    # normaliza dados
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.date
    df["Parceiro"] = pd.to_numeric(df["Parceiro"], errors="coerce").astype("Int64")

    df = df.dropna(subset=["Parceiro", "Data"])

    # cria chave de comparação
    chave_loja = df["Parceiro"].astype(str) + "_" + df["Data"].astype(str)
    chave_sistema = sistema["Parceiro"].astype(str) + "_" + sistema["Data da Venda"].astype(str)

    # encontra registros que não existem no sistema
    mask = ~chave_loja.isin(chave_sistema)

    df_filtrado = df.loc[mask]

    # retorna a linha completa (todas as colunas)
    return pd.concat([lista, df_filtrado], ignore_index=True)


# ===============================
# PROCESSAMENTO
# ===============================

diferencas = pd.DataFrame()

for nome, df_loja in dfs.items():
    diferencas = get_diferenca(df_loja, df_sistema, diferencas)


diferencas = diferencas.dropna(axis=1, how='all')


# ===============================
# RESULTADO FINAL
# ===============================

try:

    df = diferencas.dropna(subset=["Parceiro", "Data", "Dif"])
    df["Data"] = pd.to_datetime(df["Data"]).dt.date

except:
    df = diferencas









