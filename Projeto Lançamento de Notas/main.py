import threading
import time

from database import (
    criar_tabelas,inserir_devolucao,
    listar_devolucoes, buscar_ultima_devolucao,
    excluir_devolucao)

from automation import consultar_nota

import webview

from flask import (
    Flask, render_template,
    request,redirect,
    url_for,flash
)

from database import (
    criar_tabelas,
    inserir_devolucao,
    listar_devolucoes
)


# ============================================================
# CONFIGURAÇÃO DO FLASK
# ============================================================

app = Flask(__name__)
app.secret_key = "projeto-lancamento-notas"


# ============================================================
# ROTAS
# ============================================================

@app.route("/")
def index():
    """
    Tela principal do sistema.
    """
    return render_template("index.html")


# ============================================================
# DEVOLUÇÃO
# ============================================================

@app.route("/devolucao", methods=["GET", "POST"])
def devolucao():
    """
    Tela de lançamento de nota de devolução.
    """
    if request.method == "POST":

        numero_nota = request.form.get("numero_nota")
        motivo = request.form.get("motivo")
        tipo = request.form.get("tipo")

        # ----------------------------------------------------
        # VALIDAÇÃO
        # ----------------------------------------------------

        if not numero_nota or not motivo or not tipo:

            flash("Preencha todos os campos.","erro")
            return redirect(url_for("devolucao"))

        # ----------------------------------------------------
        # SALVAR NO BANCO
        # ----------------------------------------------------

        id_devolucao = inserir_devolucao(
            numero_nota=numero_nota,
            motivo=motivo,
            tipo=tipo
        )

        flash(f"Devolução cadastrada com sucesso. ID: {id_devolucao}","sucesso")

        return redirect(url_for("devolucao"))

    # --------------------------------------------------------
    # GET
    # --------------------------------------------------------

    devolucoes = listar_devolucoes()

    return render_template(
        "devolucao.html",
        devolucoes=devolucoes
    )


# ============================================================
# EXCLUIR DEVOLUÇÃO
# ============================================================

@app.route("/devolucao/excluir/<int:id_devolucao>", methods=["POST"])
def excluir_devolucao_rota(id_devolucao):
    """
    Exclui uma devolução cadastrada.
    """

    excluido = excluir_devolucao(id_devolucao)

    if excluido: flash("Devolução excluída com sucesso.","sucesso")

    else: flash("Devolução não encontrada.","erro")

    return redirect(url_for("devolucao"))

# ============================================================
# EXECUTAR AUTOMAÇÃO
# ============================================================

@app.route("/devolucao/executar", methods=["POST"])
def executar_devolucao():

    devolucao = buscar_ultima_devolucao()

    if not devolucao:
        flash("Nenhuma devolução foi cadastrada.","erro")

        return redirect(url_for("devolucao"))

    numero_nota = devolucao["numero_nota"]
    resultado = consultar_nota(numero_nota)


    if resultado["sucesso"]:
        flash(resultado["mensagem"],"sucesso")

    else:
        flash(resultado["mensagem"],"erro")

    return redirect(url_for("devolucao"))


# ============================================================
# INICIAR FLASK
# ============================================================

def iniciar_flask():
    """
    Inicia o servidor Flask localmente.
    """
    app.run(
        host="127.0.0.1",
        port=5001,
        debug=False,
        use_reloader=False
    )


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

if __name__ == "__main__":
    criar_tabelas()
    thread_flask = threading.Thread(
        target=iniciar_flask,
        daemon=True)
    thread_flask.start()
    time.sleep(1)


    webview.create_window(
        title="Projeto de Lançamento de notas",
        url="http://127.0.0.1:5001",
        width=1000,
        height=700,
        resizable=True
    )

    webview.start()