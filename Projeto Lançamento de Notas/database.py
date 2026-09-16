import sqlite3
from pathlib import Path

# ============================================================

# CONFIGURAÇÃO DO BANCO

# ============================================================

BASE_DIR = Path(__file__).resolve().parent

PASTA_DATABASE = BASE_DIR / "database"

ARQUIVO_DATABASE = PASTA_DATABASE / "lancamento_notas.db"

def conectar():

    PASTA_DATABASE.mkdir(
        parents=True,
        exist_ok=True
    )

    conexao = sqlite3.connect(ARQUIVO_DATABASE)
    conexao.row_factory = sqlite3.Row

    return conexao


# ============================================================

# CRIAÇÃO DAS TABELAS

# ============================================================

def criar_tabelas():

    conexao = conectar()
    cursor = conexao.cursor()

    # --------------------------------------------------------
    # TABELA DE DEVOLUÇÕES
    # --------------------------------------------------------
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS devolucoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data DATE NOT NULL,
            numero_nota INTEGER NOT NULL,
            motivo INTEGER NOT NULL,
            tipo INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'PENDENTE',
            data_lancamento DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    # --------------------------------------------------------
    # COMPATIBILIDADE COM BANCO EXISTENTE
    # --------------------------------------------------------
    
    cursor.execute("""
        PRAGMA table_info(devolucoes)
    """)
    
    colunas = [
        coluna["name"]
        for coluna in cursor.fetchall()
    ]
    
    if "status" not in colunas:
    
        cursor.execute("""
            ALTER TABLE devolucoes
            ADD COLUMN status TEXT
        """)
    
        cursor.execute("""
            UPDATE devolucoes
            SET status = 'PENDENTE'
            WHERE status IS NULL
        """)
    
    conexao.commit()
    conexao.close()

# ============================================================

# INSERIR DEVOLUÇÃO

# ============================================================

def inserir_devolucao(
data,
numero_nota,
motivo,
tipo
):

    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("""
        INSERT INTO devolucoes (
            data,
            numero_nota,
            motivo,
            tipo,
            status
        )
        VALUES (?, ?, ?, ?, ?)
    """, (
        data,
        numero_nota,
        motivo,
        tipo,
        "PENDENTE"
    ))
    
    conexao.commit()
    
    id_devolucao = cursor.lastrowid
    
    conexao.close()
    
    return id_devolucao

# ============================================================

# BUSCAR DEVOLUÇÃO POR ID

# ============================================================

def buscar_devolucao(id_devolucao):

    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("""
        SELECT
            id,
            data,
            numero_nota,
            motivo,
            tipo,
            status,
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


    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("""
        SELECT
            id,
            data,
            numero_nota,
            motivo,
            tipo,
            status,
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

# LISTAR DEVOLUÇÕES PENDENTES

# ============================================================

def listar_devolucoes_pendentes():

    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("""
        SELECT
            id,
            data,
            numero_nota,
            motivo,
            tipo,
            status,
            data_lancamento
        FROM devolucoes
        WHERE status = 'PENDENTE'
        ORDER BY id ASC
    """)
    
    devolucoes = cursor.fetchall()
    
    conexao.close()
    
    return [
        dict(devolucao)
        for devolucao in devolucoes
    ]

# ============================================================

# ATUALIZAR STATUS

# ============================================================

def atualizar_status_devolucao(
id_devolucao,
status
):

    status_permitidos = {
        "PENDENTE",
        "PRONTA PARA LANÇAMENTO"
    }
    
    if status not in status_permitidos:
    
        raise ValueError(
            f"Status inválido: {status}"
        )
    
    conexao = conectar()
    
    cursor = conexao.cursor()
    
    cursor.execute("""
        UPDATE devolucoes
        SET status = ?
        WHERE id = ?
    """, (
        status,
        id_devolucao
    ))
    
    conexao.commit()
    
    quantidade = cursor.rowcount
    
    conexao.close()
    
    return quantidade > 0

# ============================================================

# BUSCAR ÚLTIMA DEVOLUÇÃO PENDENTE

# ============================================================

def buscar_ultima_devolucao():

    conexao = conectar()
    cursor = conexao.cursor()
    
    cursor.execute("""
        SELECT
            id,
            data,
            numero_nota,
            motivo,
            tipo,
            status,
            data_lancamento
        FROM devolucoes
        WHERE status = 'PENDENTE'
        ORDER BY id DESC
        LIMIT 1
    """)
    
    devolucao = cursor.fetchone()
    
    conexao.close()
    
    if devolucao:
    
        return dict(devolucao)
    
    return None

# ============================================================

# EXCLUIR DEVOLUÇÃO

# ============================================================

def excluir_devolucao(id_devolucao):

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

# INICIALIZAÇÃO DO BANCO

# ============================================================

if __name__ == "__main__":

    criar_tabelas()
    print("Banco de dados inicializado com sucesso.")
    print(f"Arquivo: {ARQUIVO_DATABASE}")

