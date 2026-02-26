import sqlite3

conn = sqlite3.connect("banco.db")
c = conn.cursor()

# ✅ Corrigir programação id=2 para ter data de hoje
c.execute("UPDATE programacoes SET data = date('now') WHERE id = 2")

# ✅ Corrigir programação id=1 (data no formato correto + motorista vinculado)
c.execute("UPDATE programacoes SET motorista_id = 1, data = '2026-01-01' WHERE id = 1")

conn.commit()

c.execute("SELECT id, codigo, data, motorista_id, status FROM programacoes ORDER BY id DESC LIMIT 10")
print(c.fetchall())

conn.close()

print("✅ OK - Banco corrigido")
