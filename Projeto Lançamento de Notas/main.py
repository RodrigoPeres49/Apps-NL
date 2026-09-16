import threading
import time
from controllers.devolucao import devolucao_bp
import webview
from flask import (Flask, render_template,)
from database import (criar_tabelas)


# ============================================================
# CONFIGURAÇÃO DO FLASK
# ============================================================

app = Flask(__name__)
app.secret_key = "projeto-lancamento-notas"

app.register_blueprint(devolucao_bp,url_prefix="/devolucao")


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