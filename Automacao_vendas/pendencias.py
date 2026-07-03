import os
import pygetwindow as gw
import ctypes
from pendencias_lista import pendencias

ctypes.windll.kernel32.SetConsoleTitleW("Pendencias")
pendencias_janela = next((w for w in gw.getAllWindows() if "Pendencias" in w.title), None)
pendencias_janela.activate()

# FUNÇÃO PARA VOLTAR A JANELA PRINCIPAL

def posicionar_jan_pricipal():

    janela = next((w for w in gw.getAllWindows() if "Controle Vendas Terceirizadas NL" in w.title), None)
    janela.activate()



os.system("cls")

print("COPIE E COLE AS INFORMAÇÕES PARA ENVIAR:\n\n")

print("LISTA DE PENDENCIAS VENDAS TERCEIRIZADAS\n")

redes = set([p["Rede"] for p in pendencias])

for rede in redes:
    print(f" - {rede}")
    for p in pendencias:
        if p["Rede"] == rede:
            print(f"   {p['Dias Pendentes']}")
    print("\n")

gerar_envios = input("Deseja fazer um modelo de envio? \n- Digite S para gerar o modelo \n- Ou pressione qualquer tecla para continuar\n- ")
os.system("cls")
if gerar_envios == "S" or gerar_envios == "s":
    print("Gerando modelos de envio... \n\n\n ")
    for rede in redes: 
        print("-"*85,"\n\n")
        print("[PEDIDO] Arquivos de Vendas Pendentes\n\nSolicito no dia de hoje os arquivos de vendas diários das respectivas datas:\n")
        print(f" - {rede}")
        for p in pendencias:
            if p["Rede"] == rede:
                print(f"   {p['Dias Pendentes']}")
        print("\n")
        print("Favor enviar quando tiver possibilidade.\n\n\n")

    os.system("pause")
posicionar_jan_pricipal()

