import sqlite3
from pathlib import Path


# ============================================================
# CONFIGURAÇÃO DO BANCO
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

PASTA_DATABASE = BASE_DIR / "database"

ARQUIVO_DATABASE = PASTA_DATABASE / "lancamento_notas.db"

def conectar():
    """
    Cria e retorna uma conexão com o banco SQLite.
    """

    PASTA_DATABASE.mkdir(parents=True,exist_ok=True)
    conexao = sqlite3.connect(ARQUIVO_DATABASE)
    conexao.row_factory = sqlite3.Row

    return conexao


# ============================================================
# CRIAÇÃO DAS TABELAS
# ============================================================

def criar_tabelas():
    """
    Cria as tabelas necessárias para o sistema.
    """
    conexao = conectar()
    cursor = conexao.cursor()

    # --------------------------------------------------------
    # TABELA DE DEVOLUÇÕES
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS devolucoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_nota INTEGER NOT NULL,
            motivo INTEGER NOT NULL,
            tipo INTEGER NOT NULL,
            data_lancamento DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conexao.commit()
    conexao.close()


# ============================================================
# INSERIR DEVOLUÇÃO
# ============================================================

def inserir_devolucao(
    numero_nota,
    motivo,
    tipo
):
    """
    Insere uma nova devolução no banco.
    Retorna o ID do registro criado.
    """

    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
        INSERT INTO devolucoes (
            numero_nota,
            motivo,
            tipo
        )
        VALUES (?, ?, ?)
    """, (
        numero_nota,
        motivo,
        tipo
    ))

    conexao.commit()
    id_devolucao = cursor.lastrowid
    conexao.close()
    return id_devolucao


# ============================================================
# BUSCAR DEVOLUÇÃO POR ID
# ============================================================

def buscar_devolucao(id_devolucao):
    """
    Busca uma devolução pelo ID.
    """
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
        SELECT
            id,
            numero_nota,
            motivo,
            tipo,
            data_lancamento
        FROM devolucoes
        WHERE id = ?
    """, (
        id_devolucao,
    ))
    devolucao = cursor.fetchone()
    conexao.close()
    if devolucao:
        return dict(devolucao)
    return None


# ============================================================
# LISTAR DEVOLUÇÕES
# ============================================================

def listar_devolucoes():
    """
    Retorna todas as devoluções cadastradas.
    """
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
        SELECT
            id,
            numero_nota,
            motivo,
            tipo,
            data_lancamento
        FROM devolucoes
        ORDER BY id DESC
    """)

    devolucoes = cursor.fetchall()
    conexao.close()
    return [
        dict(devolucao)
        for devolucao in devolucoes
    ]


# ============================================================
# EXCLUIR DEVOLUÇÃO
# ============================================================

def excluir_devolucao(id_devolucao):
    """
    Exclui uma devolução pelo ID.
    """
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute("""
        DELETE FROM devolucoes
        WHERE id = ?
    """, (
        id_devolucao,
    ))
    conexao.commit()
    quantidade = cursor.rowcount
    conexao.close()
    return quantidade > 0


# ============================================================
# BUSCAR ÚLTIMA DEVOLUÇÃO
# ============================================================

def buscar_ultima_devolucao():

    conexao = conectar()

    cursor = conexao.cursor()

    cursor.execute("""
        SELECT
            id,
            numero_nota,
            motivo,
            tipo,
            data_lancamento
        FROM devolucoes
        ORDER BY id DESC
        LIMIT 1
    """)

    devolucao = cursor.fetchone()

    conexao.close()

    if devolucao:

        return dict(devolucao)

    return None

# ============================================================
# INICIALIZAÇÃO DO BANCO
# ============================================================

if __name__ == "__main__":

    criar_tabelas()
    print("Banco de dados inicializado com sucesso.")
    print(f"Arquivo: {ARQUIVO_DATABASE}")