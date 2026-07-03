import subprocess
import os

# CAMINHO DA PASTA

caminho_outros = r"C:\Users\rodrigo.souza\nlfrutas.com.br\NL_File-Server - Documentos\Comercial_Terceirizadas\CONSOLIDA_VENDAS\Vendas\Automacao_vendas"

# ARQUIVOS PARA EXECUTAR

exec_arquivos = {
    "pendencias": "pendencias.py",
    "lancamentos_Automaticos" : "lancamentos_automaticos.py",
    "relatorios" : "relatorios_csv.py"
}

# FUNCAO PARA EXECUTAR OS ARQUIVOS

def executar_externos(caminho_outros, arquivo):
    arq_externo = os.path.join(caminho_outros, arquivo)
    subprocess.run(["python", arq_externo])
