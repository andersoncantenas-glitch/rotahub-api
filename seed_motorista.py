import os
import sqlite3
import base64
import hashlib
import hmac

DB = "banco.db"

PBKDF2_ITERATIONS = 200_000

def hash_password_pbkdf2(password: str, *, iterations: int = PBKDF2_ITERATIONS) -> str:
    password = str(password or "")
    if password == "":
        raise ValueError("Senha vazia.")
    salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations, dklen=32)
    return "pbkdf2_sha256${}${}${}".format(
        iterations,
        base64.b64encode(salt).decode("ascii"),
        base64.b64encode(dk).decode("ascii"),
    )

def criar_motorista(codigo: str, nome: str, senha: str):
    conn = sqlite3.connect(DB)
    cursor = conn.cursor()

    # Verifica se já existe
    cursor.execute("SELECT id FROM motoristas WHERE UPPER(codigo)=UPPER(?) LIMIT 1", (codigo,))
    existe = cursor.fetchone()

    if existe:
        print(f"Motorista {codigo} já existe ✅ (id={existe[0]})")
        conn.close()
        return

    senha_hash = hash_password_pbkdf2(senha)
    cursor.execute(
        "INSERT INTO motoristas (codigo, nome, senha) VALUES (?, ?, ?)",
        (codigo, nome, senha_hash),
    )
    conn.commit()
    conn.close()
    print(f"Motorista {codigo} criado ✅")

if __name__ == "__main__":
    criar_motorista("MT001", "Motorista Teste", "1234")
