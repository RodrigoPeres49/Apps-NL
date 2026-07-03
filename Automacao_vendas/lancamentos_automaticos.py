import pandas as pd
import pyautogui
import time
import os
from base import *
import pygetwindow as gw
import ctypes
import win32api

ctypes.windll.kernel32.SetConsoleTitleW("Lançamentos Vendas")

# POSICIONANDO JANELAS PARA REALIZAR O LANÇAMENTO


def posiciona_lancamentos():
    lancamentos = next((w for w in gw.getAllWindows() if "Lançamentos Vendas" in w.title), None)
    
    if not lancamentos:
        print("Janela do Lançamentos não encontrada!")
        return
    
    lancamentos.activate()
    time.sleep(1.5)
    pyautogui.hotkey('win', 'up')
    pyautogui.hotkey('win', 'up')
    pyautogui.press('esc')
    time.sleep(0.5)
    pyautogui.press('esc')
    pyautogui.press('esc')
    pyautogui.hotkey('win', 'up')
    pyautogui.hotkey('win', 'right')
    pyautogui.hotkey('win', 'right')



def posicionar_sankhya():
    sankhya = next((w for w in gw.getAllWindows() if "Sankhya" in w.title), None)

    if not sankhya:
        print("Janela do Sankhya não encontrada!")
        return

    sankhya.activate()
    time.sleep(1.5)
    pyautogui.hotkey('win', 'left') 
    pyautogui.hotkey('win', 'up')
    pyautogui.hotkey('win', 'up')
    time.sleep(0.5)
    pyautogui.hotkey('win', 'up')
    time.sleep(0.5)
    pyautogui.hotkey('win', 'left')

# POSICIONAR JANELA PRINCIPAL

def posicionar_jan_pricipal():

    janela = next((w for w in gw.getAllWindows() if "Controle Vendas Terceirizadas NL" in w.title), None)
    janela.activate()


# POSICIONAR JANELAS

posiciona_lancamentos()
posicionar_sankhya()


# INICIANDO LANÇAMENTOS

os.system("cls")
print(""" ------------------ APLICAÇÃO PARA LANÇAR VENDAS AUTOMATICAMENTE ------------------
      - Abra o Sankhya, \n
      - Vá em Registro de Vendas Terceirizadas,\n
      - Clique em novo Registro, \n
      - Posicione a Janela do Sankhya ao lado da janela deste programa \n 
       ( Você pode segurar a tecla: \n Windows e apertar a seta para um dos lados para posicionar )\n
      - Antes de iniciar,\n Verifique se todas as planilhas foram atualizadas para evitar erros no procedimento \n
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
    "eleve": os.path.join(base_dir, "imagens", "Eleve.png"),
    "salvar": os.path.join(base_dir, "imagens", "Salvar.png"),
    "adicionar": os.path.join(base_dir, "imagens","Adicionar.png"),
    "descartar": os.path.join(base_dir, "imagens","Descartar.png")
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

# SELECIONAR O SETOR ESPECÍFICO

def selecionar_setor(valor_setor):
    if not pyautogui.locateCenterOnScreen(imgs["setor"], confidence=0.8): # estava 0.5
        print("Campo Setor da Loja não encontrado.")
        return
    pyautogui.click(pyautogui.locateCenterOnScreen(imgs["setor"], confidence=0.8)) # estava 0.766
    time.sleep(0.2)
    if valor_setor.upper() == "LOJA":
        img_setor = imgs["flv"]
    elif valor_setor.upper() == "FLORES":
        img_setor = imgs["flores"]
    else:
        img_setor =imgs["eleve"]
    
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
    salvar = pyautogui.locateCenterOnScreen(imgs["salvar"], confidence=0.9)

    if salvar:
        pyautogui.click(salvar)
        time.sleep(1.5)
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

# FUNÇÃO PARA FAZER OUTRA TENTATIVA DE LANÇAMENTO

def outra_tentativa():

    descartar = pyautogui.locateCenterOnScreen(imgs["descartar"], confidence=0.9)
    time.sleep(5)
    pyautogui.click(descartar)
    adicionar = pyautogui.locateCenterOnScreen(imgs["adicionar"], confidence=0.9)
    time.sleep(1)
    pyautogui.click(adicionar)
    time.sleep(1)


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
        
        print("\nErro ao realizar lançamento, vamos tentar realizar o lançamento novamente\n")
        
        try:

            # TENTANDO NOVAMENTE SE ERRO

            outra_tentativa()
            outra_tentativa()

            preencher_formulario(parceiro, data_venda, "LOJA", valor_loja, qtd_loja)
            print(f"Venda: {parceiro}, {data_venda}, Setor: LOJA, Valor: {valor_loja}, Qtd:{qtd_loja} Lançado com sucesso!")

        except Exception as e:
            
            print(f"\nOcorreu um erro ao tentar preencher a Venda:\n {parceiro}\n {data_venda}\n Setor: LOJA \n Valor: {valor_loja}\n Qtd:{qtd_loja}\n Mensagem de Erro: {e}\n Favor verificar")
            print("\nAntes de prosseguir ou realizar algum lançamento,\n certifique que o formulário esteja vazio para evitar erros no prossedimento")
            print("\nCaso for feito algum lançamento salve a informação,\nclique em ' + ' antes de dar continuidade a automação \n")
            os.system("pause")
    
    # LANÇAMENTO FLORES

    try:

        if pd.notna(row.get("FLORES")) and pd.notna(row.get("Qtd_FLORES")):
            valor_flores = f'{float(row["FLORES"]):.2f}'.replace('.', ',')
            qtd_flores = f'{float(row["Qtd_FLORES"]):.2f}'.replace('.', ',')
            preencher_formulario(parceiro, data_venda, "FLORES", valor_flores, qtd_flores)
            print(f"Venda: {parceiro}, {data_venda}, Setor: FLORES, Valor: {valor_flores}, Qtd:{qtd_flores} Lançado com sucesso!")
    
    except:

        print("\nErro ao realizar lançamento, vamos tentar realizar o lançamento novamente\n")
        
        try:

            # TENTANDO NOVAMENTE SE ERRO

            outra_tentativa()
            outra_tentativa()

            preencher_formulario(parceiro, data_venda, "FLORES", valor_flores, qtd_flores)
            print(f"Venda: {parceiro}, {data_venda}, Setor: FLORES, Valor: {valor_flores}, Qtd:{qtd_flores} Lançado com sucesso!")

        except Exception as e:

            print(f"\nOcorreu um erro ao tentar preencher a Venda:\n {parceiro}\n {data_venda}\n Setor: FLORES \n Valor: {valor_flores}\n Qtd:{qtd_flores}\n Mensagem de Erro: {e} \n Favor verificar")
            print("\n Antes de prosseguir ou realizar algum lançamento,\n certifique que o formulário esteja vazio para evitar erros no prossedimento")
            print("\nCaso for feito algum lançamento salve a informação,\nclique em ' + ' antes de dar continuidade a automação \n")
            os.system("pause")

    #LANÇAMENTO ELEVE

    try:

       if pd.notna(row.get("ELEVE")) and pd.notna(row.get("Qtd_ELEVE")):
           valor_eleve = f'{float(row["ELEVE"]):.2f}'.replace('.', ',')
           qtd_eleve = f'{float(row["Qtd_ELEVE"]):.2f}'.replace('.', ',')
           preencher_formulario(parceiro, data_venda, "ELEVE", valor_eleve, qtd_eleve)
           print(f"Venda: {parceiro}, {data_venda}, Setor: ELEVE, Valor: {valor_eleve}, Qtd:{qtd_eleve} Lançado com sucesso!")
    
    except:

        print("\nErro ao realizar lançamento, vamos tentar realizar o lançamento novamente\n")

        try:

            # TENTANDO NOVAMENTE SE ERRO

            outra_tentativa()
            outra_tentativa()


            preencher_formulario(parceiro, data_venda, "ELEVE", valor_eleve, qtd_eleve)
            print(f"Venda: {parceiro}, {data_venda}, Setor: ELEVE, Valor: {valor_eleve}, Qtd:{qtd_eleve} Lançado com sucesso!")

        except Exception as e:

            print(f"\nOcorreu um erro ao tentar preencher a Venda:\n {parceiro} \n{data_venda}\n Setor: ELEVE \n Valor: {valor_eleve}\n Qtd:{qtd_eleve}\n Mensagem de erro: {e}\n Favor verificar")
            print("\n Antes de prosseguir ou realizar algum lançamento,\n certifique que o formulário esteja vazio para evitar erros no prossedimento")
            print("\nCaso for feito algum lançamento salve a informação,\nclique em ' + ' antes de dar continuidade a automação \n")
            os.system("pause")


print("Lançamentos concluídos com sucesso!")
os.system("pause")
posicionar_jan_pricipal()
