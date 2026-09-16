from flask import (Blueprint,render_template,request,redirect,url_for,flash)
from database import (inserir_devolucao,listar_devolucoes_pendentes,listar_devolucoes,excluir_devolucao)
from automations.automation_devolucao import consultar_nota_devolucao
from database import (inserir_devolucao,listar_devolucoes,listar_devolucoes_pendentes, excluir_devolucao)


devolucao_bp = Blueprint("devolucao",__name__)

# ============================================================
# DEVOLUÇÃO
# ============================================================

@devolucao_bp.route("/", methods=["GET", "POST"])
def devolucao():


    if request.method == "POST":

        data = request.form.get("data")
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
            data = data,
            numero_nota=numero_nota,
            motivo=motivo,
            tipo=tipo
        )

        flash(f"Devolução cadastrada com sucesso. ID: {id_devolucao}","sucesso")

        return redirect(url_for("devolucao.devolucao"))

    # --------------------------------------------------------
    # GET
    # --------------------------------------------------------

    devolucoes = listar_devolucoes_pendentes()

    return render_template(
        "devolucao.html",
        devolucoes=devolucoes
    )


# ============================================================
# EXCLUIR DEVOLUÇÃO
# ============================================================

@devolucao_bp.route("/devolucao/excluir/<int:id_devolucao>", methods=["POST"])
def excluir_devolucao_rota(id_devolucao):

    excluido = excluir_devolucao(id_devolucao)

    if excluido: flash("Devolução excluída com sucesso.","sucesso")

    else: flash("Devolução não encontrada.","erro")

    return redirect(url_for("devolucao.devolucao"))

# ============================================================
# EXECUTAR AUTOMAÇÃO DEVOLUÇÃO
# ============================================================

@devolucao_bp.route("/devolucao/executar", methods=["POST"])
def executar_devolucao():

    devolucao = listar_devolucoes()

    for nf in devolucao: 
        if not nf:
            flash("Nenhuma devolução foi cadastrada.","erro")
    
            return redirect(url_for("devolucao"))
    
        # VARIÁVEIS DA NOTA
        
        numero_nota = nf["numero_nota"]
        data = nf["data"]
        tipo = nf["tipo"]
        motivo = nf["motivo"]
        resultado = consultar_nota_devolucao(numero_nota, data, tipo, motivo)


        if resultado["sucesso"]:
            flash(resultado["mensagem"],"sucesso")
    
        else:
            flash(resultado["mensagem"],"erro")

    return redirect(url_for("devolucao.devolucao"))