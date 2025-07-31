import pandas as pd 
import os
from datetime import datetime
import json

ARQUIVO = "dados.xlsx"
LOJA_JSON = "static/nome_loja.json"

class Api:

    # INICIANDO API

    def __init__(self):
        if not os.path.exists(ARQUIVO):
            print("Criando planilha...")
            df = pd.DataFrame(columns=["idRegistro","Data", "Produto", "Quantidade", "Preço", "Valor Total"])
            df.to_excel(ARQUIVO, index=False)


    # CARREGANDO ARQUIVO COM O NOME DA LOJA EM JSON

    def carregar_nome_loja(self):
        try:
            with open(LOJA_JSON, "r") as f:
                arquivo_loja = json.load(f)
            return arquivo_loja.get("loja", "")
        except:
            return ""
    
    # ATUALIZANDO OU INSERINDO O NOME DA LOJA 

    def definir_nome_loja(self,nome): 
        try:
            with open(LOJA_JSON, "r") as f:
                arquivo_loja = json.load(f)
            
            arquivo_loja["loja"] = nome
    
            with open(LOJA_JSON, "w") as f:
                json.dump(arquivo_loja, f, indent=4)
            
            print(arquivo_loja)

        except Exception as e:
            print(f"Ocorreu um erro ao atualizar o arquivo JSON: {e}")



    # ADICIONANDO ITEM

    def adicionar_item(self, data, produto, quantidade, preco, valor):

        if os.path.exists(ARQUIVO):
            df = pd.read_excel(ARQUIVO)
        else:
            df = pd.DataFrame(columns=["idRegistro", "Data", "Produto", "Quantidade", "Preço", "Valor Total"])
    
        novo_id = int(df["idRegistro"].max()) + 1 if not df.empty else 1
        id = novo_id
    
        nova_linha = {
            "idRegistro": id,
            "Data": data,
            "Produto": produto,
            "Quantidade": str(quantidade).replace(".", ","),
            "Preço": str(preco).replace(".", ","),
            "Valor Total": str(valor).replace(".", ",")
        }
        

        df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
        df.to_excel(ARQUIVO, index=False)
    
        return df.to_dict(orient="records")
    
    # APAGANDO ITEM
    
    def apagar_item(self, idRegistro):
        df = pd.read_excel(ARQUIVO)
        idRegistro = int(idRegistro)
        df = df[df["idRegistro"] != idRegistro]
        df.to_excel("dados.xlsx", index=False)

    # EDITANDO ITEM
    
    def editar_item(self, idRegistro, data, produto, quantidade, preco, valor):
        df = pd.read_excel(ARQUIVO)
    
        idRegistro = int(idRegistro)
    
        index = df[df["idRegistro"] == idRegistro].index
    
        if not index.empty:
            i = index[0] 
    
            df.at[i, "Data"] = data
            df.at[i, "Produto"] = produto
            df.at[i, "Quantidade"] = str(quantidade).replace(".", ",")
            df.at[i, "Preço"] = str(preco).replace(".", ",")
            df.at[i, "Valor Total"] = str(valor).replace(".", ",")
    
            df.to_excel(ARQUIVO, index=False)
            print(f" Linha editada: idRegistro {idRegistro}, índice real {i}")
        else:
            print(f" ID {idRegistro} não encontrado")

    
    # ARMAZENAR DADOS NA PASTA ARQUIVOS E LIMPAR PLANILHA DE DADOS

    def armazenar_planilha(self):
        df = pd.read_excel(ARQUIVO)

        if not os.path.exists("arquivos"):
            os.makedirs("arquivos")
        
    
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        caminho = f"arquivos/{self.carregar_nome_loja()}_consinterno_{timestamp}.xlsx"
    
        df.to_excel(caminho, index=False)
    
        df_vazio = pd.DataFrame(columns=["idRegistro", "Data", "Produto", "Quantidade", "Preço", "Valor Total"])
        df_vazio.to_excel(ARQUIVO, index=False)
    
        print(f" Arquivo salvo em: {caminho}")
        return caminho
    
    # CARREGAR LOJAS

    def carregar_lojas(self):
        with open("static/lojas.json") as f:
            return json.load(f)
        
    # CARREGANDO DADOS

    def carregar_dados(self):
        df = pd.read_excel(ARQUIVO)
        return df.to_dict(orient="records")
    