import pandas as pd
import sqlite3

def importar_excel(caminho_arquivo):
    # Ler o arquivo Excel
    df = pd.read_excel(caminho_arquivo)

    # Conectar ao banco
    conn = sqlite3.connect("banco.db")
    cursor = conn.cursor()

    # Percorrer cada linha do Excel
    for index, row in df.iterrows():
        nome = row["Cliente"]
        telefone = str(row["Telefone"])
        endereco = row["Endereco"]

        cursor.execute("""
        INSERT INTO clientes (nome, telefone, endereco)
        VALUES (?, ?, ?)
        """, (nome, telefone, endereco))

    # Salvar e fechar
    conn.commit()
    conn.close()

    print("Importação concluída com sucesso!")
