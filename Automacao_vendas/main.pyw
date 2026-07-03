import webview
import threading
from base import *
from pendencias_lista import pendencias
from externos import *
from flask import Flask, render_template, redirect

app = Flask(__name__)

@app.route("/", methods=["GET"])
def index():
    return render_template(
        "index.html",
        arquivos=arquivos,
        caminho=caminho,
        pendencias =pendencias,
        manual = manual,
        controle_pendencias = controle_pendencias
    )

@app.route("/abrir/<string:arquivo>")
def abrir(arquivo):
    abrir_arquivo(caminho, arquivo)  
    return redirect("/")

@app.route("/abrirmanual/<string:arquivo>")
def abrir_manual(arquivo):
    abrir_arquivo(caminho_manual,arquivo)
    return redirect("/")

@app.route("/abrir_outros/<string:arquivo>")
def abrir_outros(arquivo):
    executar_externos(caminho_outros,exec_arquivos[arquivo])
    return redirect("/")


def start_flask():
    app.run(host="127.0.0.1", port=5000)

if __name__ == "__main__":


    threading.Thread(target=start_flask, daemon=True).start()
    
    janela_main = webview.create_window("Controle Vendas Terceirizadas NL", "http://127.0.0.1:5000", maximized = True)
    webview.start()
    
