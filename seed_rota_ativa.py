import sqlite3
from datetime import datetime
import os
import sys

# =========================================================
# CONFIG
# =========================================================
DB_PATH = os.path.abspath("banco.db")  # ✅ usa o banco.db do backend
MOTORISTA_CODIGO = "MT001"
CODIGO_PROGRAMACAO = "PGTESTE001"
STATUS = "ATIVA"

VEICULO = "AAA9A00"
EQUIPE = "EQ1"
TIPO_ROTA = "SERRA"
TOTAL_CAIXAS = 10

print("======================================")
print("✅ SEED ROTA ATIVA")
print("DB:", DB_PATH)
print("======================================")

if not os.path.exists(DB_PATH):
    print("❌ banco.db não encontrado na pasta atual.")
    print("➡️ Rode este script dentro de C:\\pdc_rota")
    sys.exit(1)

conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()

# =========================================================
# 1) Buscar motorista
# =========================================================
cur.execute(
    "SELECT id, nome FROM motoristas WHERE UPPER(codigo)=UPPER(?) LIMIT 1",
    (MOTORISTA_CODIGO,)
)
m = cur.fetchone()

if not m:
    print(f"❌ Motorista {MOTORISTA_CODIGO} não encontrado.")
    conn.close()
    sys.exit(1)

motorista_id = m[0]
motorista_nome = m[1]

print("✅ Motorista:", motorista_nome, "| id:", motorista_id)

# =========================================================
# 2) Verificar se já existe programação
# =========================================================
cur.execute(
    "SELECT id, codigo_programacao, status FROM programacoes WHERE codigo_programacao=? LIMIT 1",
    (CODIGO_PROGRAMACAO,)
)
ex = cur.fetchone()

if ex:
    print("⚠️ Já existe programação:", ex)
    conn.close()
    sys.exit(0)

# =========================================================
# 3) Inserir programação ATIVA
# =========================================================
dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

sql = """
INSERT INTO programacoes (
    codigo_programacao,
    codigo,
    motorista_id,
    motorista,
    veiculo,
    equipe,
    tipo_rota,
    status,
    data,
    data_criacao,
    total_caixas
) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""

params = (
    CODIGO_PROGRAMACAO,
    CODIGO_PROGRAMACAO,
    motorista_id,
    motorista_nome,
    VEICULO,
    EQUIPE,
    TIPO_ROTA,
    STATUS,
    dt,
    dt,
    TOTAL_CAIXAS,
)

cur.execute(sql, params)

conn.commit()
conn.close()

print("======================================")
print("✅ ROTA ATIVA CRIADA COM SUCESSO:", CODIGO_PROGRAMACAO)
print("➡️ Agora rode /rotas/ativas e atualize o app.")
print("======================================")
