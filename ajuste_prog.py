import sqlite3

conn = sqlite3.connect("banco.db")
cursor = conn.cursor()

# APAGA tabela antiga
cursor.execute("DROP TABLE IF EXISTS programacoes")

# CRIA tabela correta e definitiva
cursor.execute("""
CREATE TABLE programacoes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT UNIQUE,
    data TEXT,
    motorista TEXT,
    veiculo TEXT,
    equipe TEXT,
    total_caixas INTEGER,
    kg_estimado REAL,
    kg_real REAL,
    status TEXT DEFAULT 'AGUARDANDO NF'
)
""")

conn.commit()
conn.close()

print("Tabela programacoes recriada com sucesso!")
