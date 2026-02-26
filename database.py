import sqlite3

DB = "banco.db"

def conectar():
    return sqlite3.connect(DB)

def _colunas_da_tabela(cursor, tabela: str) -> set[str]:
    cursor.execute(f"PRAGMA table_info({tabela})")
    return {row[1] for row in cursor.fetchall()}

def _add_coluna_se_nao_existir(cursor, tabela: str, coluna: str, ddl: str):
    cols = _colunas_da_tabela(cursor, tabela)
    if coluna not in cols:
        cursor.execute(f"ALTER TABLE {tabela} ADD COLUMN {ddl}")

def criar_banco():
    conn = conectar()
    cursor = conn.cursor()

    # ===================== USUÁRIOS =====================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        senha TEXT NOT NULL,
        permissoes TEXT DEFAULT 'OPERADOR',
        cpf TEXT,
        idade INTEGER,
        telefone TEXT
    )
    """)
    _add_coluna_se_nao_existir(cursor, "usuarios", "permissoes", "permissoes TEXT DEFAULT 'OPERADOR'")
    _add_coluna_se_nao_existir(cursor, "usuarios", "cpf", "cpf TEXT")
    _add_coluna_se_nao_existir(cursor, "usuarios", "idade", "idade INTEGER")
    _add_coluna_se_nao_existir(cursor, "usuarios", "telefone", "telefone TEXT")

    # ===================== MOTORISTAS =====================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS motoristas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT,
    nome TEXT NOT NULL,
    cnh TEXT,
    cpf TEXT,
    idade INTEGER,
    telefone TEXT,
    senha TEXT
)
""")

    # Adiciona a coluna sem UNIQUE (SQLite não permite ADD COLUMN com UNIQUE)
    _add_coluna_se_nao_existir(cursor, "motoristas", "codigo", "codigo TEXT")
    _add_coluna_se_nao_existir(cursor, "motoristas", "cnh", "cnh TEXT")
    _add_coluna_se_nao_existir(cursor, "motoristas", "cpf", "cpf TEXT")
    _add_coluna_se_nao_existir(cursor, "motoristas", "idade", "idade INTEGER")
    _add_coluna_se_nao_existir(cursor, "motoristas", "telefone", "telefone TEXT")
    _add_coluna_se_nao_existir(cursor, "motoristas", "senha", "senha TEXT")

    # Cria UNIQUE INDEX separado (isso sim é permitido)
    cursor.execute("""
    CREATE UNIQUE INDEX IF NOT EXISTS idx_motoristas_codigo_unique
    ON motoristas (codigo)
""")

    # ===================== VEÍCULOS =====================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS veiculos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        placa TEXT NOT NULL UNIQUE,
        modelo TEXT,
        ano TEXT,
        cor TEXT
    )
    """)
    _add_coluna_se_nao_existir(cursor, "veiculos", "modelo", "modelo TEXT")
    _add_coluna_se_nao_existir(cursor, "veiculos", "ano", "ano TEXT")
    _add_coluna_se_nao_existir(cursor, "veiculos", "cor", "cor TEXT")

    # ===================== EQUIPES =====================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS equipes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT NOT NULL UNIQUE,
        ajudante1 TEXT,
        ajudante2 TEXT
    )
    """)
    _add_coluna_se_nao_existir(cursor, "equipes", "codigo", "codigo TEXT")
    _add_coluna_se_nao_existir(cursor, "equipes", "ajudante1", "ajudante1 TEXT")
    _add_coluna_se_nao_existir(cursor, "equipes", "ajudante2", "ajudante2 TEXT")

    # ===================== VENDEDORES =====================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS vendedores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        telefone TEXT,
        rota TEXT
    )
    """)
    _add_coluna_se_nao_existir(cursor, "vendedores", "telefone", "telefone TEXT")
    _add_coluna_se_nao_existir(cursor, "vendedores", "rota", "rota TEXT")

    # ===================== VENDAS IMPORTADAS (WIBI) =====================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS vendas_importadas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        pedido TEXT,
        data_venda TEXT,
        cliente TEXT,
        nome_cliente TEXT,
        vendedor TEXT,
        produto TEXT,
        caixas INTEGER,
        kg_cliente REAL,
        valor REAL,
        bonificacao TEXT,
        observacao TEXT,
        importado_em TEXT DEFAULT (datetime('now'))
    )
    """)
    _add_coluna_se_nao_existir(cursor, "vendas_importadas", "kg_cliente", "kg_cliente REAL")
    _add_coluna_se_nao_existir(cursor, "vendas_importadas", "produto", "produto TEXT")
    _add_coluna_se_nao_existir(cursor, "vendas_importadas", "bonificacao", "bonificacao TEXT")
    _add_coluna_se_nao_existir(cursor, "vendas_importadas", "observacao", "observacao TEXT")
    _add_coluna_se_nao_existir(cursor, "vendas_importadas", "importado_em", "importado_em TEXT")

    # ===================== PROGRAMAÇÕES =====================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS programacoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT UNIQUE,
        data TEXT DEFAULT (date('now')),
        motorista_id INTEGER,
        veiculo_id INTEGER,
        equipe_id INTEGER,
        total_caixas INTEGER DEFAULT 0,
        kg_estimado REAL,
        kg_real REAL,
        status TEXT DEFAULT 'AGUARDANDO NF',
        FOREIGN KEY (motorista_id) REFERENCES motoristas(id),
        FOREIGN KEY (veiculo_id) REFERENCES veiculos(id),
        FOREIGN KEY (equipe_id) REFERENCES equipes(id)
    )
    """)
    _add_coluna_se_nao_existir(cursor, "programacoes", "motorista_id", "motorista_id INTEGER")
    _add_coluna_se_nao_existir(cursor, "programacoes", "veiculo_id", "veiculo_id INTEGER")
    _add_coluna_se_nao_existir(cursor, "programacoes", "equipe_id", "equipe_id INTEGER")
    _add_coluna_se_nao_existir(cursor, "programacoes", "total_caixas", "total_caixas INTEGER DEFAULT 0")
    _add_coluna_se_nao_existir(cursor, "programacoes", "kg_estimado", "kg_estimado REAL")
    _add_coluna_se_nao_existir(cursor, "programacoes", "kg_real", "kg_real REAL")
    _add_coluna_se_nao_existir(cursor, "programacoes", "status", "status TEXT DEFAULT 'AGUARDANDO NF'")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS programacao_itens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        programacao_id INTEGER NOT NULL,
        venda_id INTEGER NOT NULL,
        caixas INTEGER,
        preco REAL,
        kg_cliente REAL,
        FOREIGN KEY (programacao_id) REFERENCES programacoes(id),
        FOREIGN KEY (venda_id) REFERENCES vendas_importadas(id)
    )
    """)
    _add_coluna_se_nao_existir(cursor, "programacao_itens", "caixas", "caixas INTEGER")
    _add_coluna_se_nao_existir(cursor, "programacao_itens", "preco", "preco REAL")
    _add_coluna_se_nao_existir(cursor, "programacao_itens", "kg_cliente", "kg_cliente REAL")

    # ===================== PDC =====================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS pdc_lancamentos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        programacao_id INTEGER NOT NULL,
        venda_id INTEGER NOT NULL,
        pago INTEGER DEFAULT 0,
        valor_pago REAL DEFAULT 0,
        forma_pagamento TEXT DEFAULT 'AVISTA',
        observacao TEXT,
        numero_nf TEXT,
        atualizado_em TEXT DEFAULT (datetime('now')),
        UNIQUE(programacao_id, venda_id),
        FOREIGN KEY (programacao_id) REFERENCES programacoes(id),
        FOREIGN KEY (venda_id) REFERENCES vendas_importadas(id)
    )
    """)

    # ===================== FECHAMENTO =====================
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS veiculo_km (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        placa TEXT NOT NULL UNIQUE,
        km_atual REAL DEFAULT 0,
        atualizado_em TEXT DEFAULT (datetime('now'))
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS fechamento_rotas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        programacao_id INTEGER NOT NULL UNIQUE,
        km_saida REAL,
        km_chegada REAL,
        litros REAL,
        media REAL,
        custo REAL,
        cx_carregada INTEGER,
        kg_nf REAL,
        aves_por_caixa INTEGER DEFAULT 6,
        kg_carregado REAL,
        saldo REAL,
        adiantamento REAL DEFAULT 0,
        devolver REAL DEFAULT 0,
        cheque REAL DEFAULT 0,
        valor_caixa REAL DEFAULT 0,
        total_dinheiro REAL DEFAULT 0,
        criado_em TEXT DEFAULT (datetime('now')),
        FOREIGN KEY (programacao_id) REFERENCES programacoes(id)
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS fechamento_despesas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        programacao_id INTEGER NOT NULL,
        descricao TEXT,
        valor REAL DEFAULT 0,
        FOREIGN KEY (programacao_id) REFERENCES programacoes(id)
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS fechamento_cedulas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        programacao_id INTEGER NOT NULL,
        valor_cedula REAL NOT NULL,
        quantidade INTEGER DEFAULT 0,
        total REAL DEFAULT 0,
        UNIQUE(programacao_id, valor_cedula),
        FOREIGN KEY (programacao_id) REFERENCES programacoes(id)
    )
    """)

    conn.commit()
    conn.close()

if __name__ == "__main__":
    criar_banco()
    print("BANCO DE DADOS CRIADO/ATUALIZADO COM SUCESSO!")

