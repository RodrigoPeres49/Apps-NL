import pandas as pd
import pyautogui
import time
import os
from base import *


# INICIANDO LANÇAMENTOS

os.system("cls")
print(""" ------------------ APLICAÇÃO PARA LANÇAR VENDAS AUTOMATICAMENTE ------------------
      - Abra o Sankhya, \n
      - Vá em Registro de Vendas Terceirizadas,\n
      - Clique em novo Registro, \n
      - Posicione a Janela do Sankhya ao lado da janela deste programa \n 
       ( Você pode segurar a tecla Windows e apertar a seta para um dos lados para posicionar )\n
      - Antes de iniciar, Verifique se todas as planilhas foram atualizadas para evitar erros no procedimento \n
           """)
print("Verifique todos os itens que irão ser lançados no preenchimento:")
print(df.head(350))

os.system("pause")


base_dir = os.path.dirname(__file__)
imgs = {
    "parceiro": os.path.join(base_dir, "imagens", "Parceiro.png"),
    "data": os.path.join(base_dir, "imagens", "Data_Venda.png"),
    "valor": os.path.join(base_dir, "imagens", "Valor_Venda.png"),
    "quantidade": os.path.join(base_dir, "imagens", "Quantidade.png"),
    "setor": os.path.join(base_dir, "imagens", "Setor_Loja.png"),
    "setor_campos": os.path.join(base_dir, "imagens", "Setor_Loja_Campos.png"), 
    "flv": os.path.join(base_dir, "imagens", "Loja.png"),
    "flores": os.path.join(base_dir, "imagens", "Flores.png"),
    "salvar": os.path.join(base_dir, "imagens", "Salvar.png"),
    "adicionar": os.path.join(base_dir, "imagens","Adicionar.png"),
}

pyautogui.PAUSE = 0.5


def localizar_e_digitar(img, texto):
    campo = pyautogui.locateCenterOnScreen(img, confidence=0.6)
    if campo:
        pyautogui.click(campo)
        pyautogui.write(texto)
        
    else:
        print(f"Campo não encontrado: {img}")
        return False
    return True

def selecionar_setor(valor_setor):
    if not pyautogui.locateCenterOnScreen(imgs["setor"], confidence=0.5):
        print("Campo Setor da Loja não encontrado.")
        return
    pyautogui.click(pyautogui.locateCenterOnScreen(imgs["setor"], confidence=0.766))
    time.sleep(0.2)
    if valor_setor.upper() == "LOJA":
        img_setor = imgs["flv"]
    else:
        img_setor = imgs["flores"]
    
    time.sleep(0.2)
    setor = pyautogui.locateCenterOnScreen(img_setor, confidence=0.95)
    if setor:
        pyautogui.click(setor)
    else:
        print(f"Opção '{valor_setor}' não encontrada no dropdown.")

# PREENCHER O FORMULARIO COM OS DADOS

def preencher_formulario(parceiro, data_venda, setor, valor, quantidade):
    if not localizar_e_digitar(imgs["parceiro"], parceiro): return
    selecionar_setor(setor)
    if not localizar_e_digitar(imgs["data"], data_venda): return
    if not localizar_e_digitar(imgs["valor"], valor): return
    if not localizar_e_digitar(imgs["quantidade"], quantidade): return

    # SALVAR
    salvar = pyautogui.locateCenterOnScreen(imgs["salvar"], confidence=0.5)
    if salvar:
        pyautogui.click(salvar)
        time.sleep(1)
        adicionar = pyautogui.locateCenterOnScreen(imgs["adicionar"], confidence=0.9)
        if adicionar:
            pyautogui.click(adicionar)
        else:
            print("Botão para adicionar uma nova venda não encontrado.")
    else:
        print("Botão de salvar não encontrado.")

    time.sleep(0.5)

try:
    df = diferencas.dropna(subset=["Parceiro", "Data"])
except:
    os.system("cls")
    print("A lista está vazia, favor atualizar as planilhas para verificar as pendências.")
    os.system("pause")

os.system("cls")
print("Iniciando preenchimento automático...")
print("Vendas lançadas:")

for idx, row in df.iterrows():
    parceiro = str(int(row["Parceiro"]))
    data_venda = pd.to_datetime(row["Data"]).strftime("%d/%m/%Y")
    
# LANÇAMENTO FLV

    try:

        if pd.notna(row.get("FLV")) and pd.notna(row.get("Qtd_FLV")):
            valor_loja = f'{float(row["FLV"]):.2f}'.replace('.', ',')
            qtd_loja = f'{float(row["Qtd_FLV"]):.2f}'.replace('.', ',')
            preencher_formulario(parceiro, data_venda, "LOJA", valor_loja, qtd_loja)
            print(f"Venda: {parceiro}, {data_venda}, Setor: LOJA, Valor: {valor_loja}, Qtd:{qtd_loja} Lançado com sucesso!")
    
    except:
        print(f"Ocorreu um erro ao tentar preencher a Venda: {parceiro}, {data_venda}, Setor: LOJA, Valor: {valor_loja}, Qtd:{qtd_loja} Favor verificar")
        os.system("pause")
    
# LANÇAMENTO FLORES

    try:

        if pd.notna(row.get("FLORES")) and pd.notna(row.get("Qtd_FLORES")):
            valor_flores = f'{float(row["FLORES"]):.2f}'.replace('.', ',')
            qtd_flores = f'{float(row["Qtd_FLORES"]):.2f}'.replace('.', ',')
            preencher_formulario(parceiro, data_venda, "FLORES", valor_flores, qtd_flores)
            print(f"Venda: {parceiro}, {data_venda}, Setor: FLORES, Valor: {valor_flores}, Qtd:{qtd_flores} Lançado com sucesso!")
    
    except:
        print(f"Ocorreu um erro ao tentar preencher a Venda:\n {parceiro}, {data_venda}, Valor: {valor_flores}, Qtd:{qtd_flores} Favor verificar")
        os.system("pause")


print("Lançamentos concluídos com sucesso!")
os.system("pause")
