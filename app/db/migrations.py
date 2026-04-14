def table_has_column(cur, table, col):
    """Verifica se coluna existe na tabela"""
    cur.execute(f"PRAGMA table_info({table})")
    cols = [r[1] for r in cur.fetchall()]
    return col in cols


def safe_add_column(cur, table, col, coltype):
    """Adiciona coluna apenas se nao existir"""
    if not table_has_column(cur, table, col):
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {coltype}")
