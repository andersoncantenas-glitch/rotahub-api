import sqlite3
import uuid

def criar_programacao():
    # Gerar código único mais compacto (6 caracteres hex)
    codigo = uuid.uuid4().hex[:6].upper()

    print("=== CRIAR PROGRAMAÇÃO DE ENTREGA ===")

    motorista = input("Nome do motorista: ")
    veiculo = input("Veículo: ")
    equipe = input("Equipe de ajudantes: ")
    kg_carregado = float(input("Quilos carregados: "))
    granja = input("Granja: ")

    # Conectar ao banco
    conn = sqlite3.connect("banco.db")
    cursor = conn.cursor()

    cursor.execute("""
    INSERT INTO programacoes
    (codigo, motorista, veiculo, equipe, kg_carregado, granja)
    VALUES (?, ?, ?, ?, ?, ?)
    """, (codigo, motorista, veiculo, equipe, kg_carregado, granja))

    conn.commit()
    conn.close()
    return codigo

    print("\nProgramação criada com sucesso!")
    print(f"Código da programação: {codigo}")
def listar_clientes():
    conn = sqlite3.connect("banco.db")
    cursor = conn.cursor()

    cursor.execute("SELECT id, nome FROM clientes")
    clientes = cursor.fetchall()

    conn.close()

    print("\nClientes cadastrados:")
    for cliente in clientes:
        print(f"{cliente[0]} - {cliente[1]}")
def adicionar_entregas(codigo_programacao):
    conn = sqlite3.connect("banco.db")
    cursor = conn.cursor()

    while True:
        listar_clientes()
        cliente_id = input("\nDigite o ID do cliente (ou ENTER para finalizar): ")

        if cliente_id == "":
            break

        caixas = int(input("Quantidade de caixas: "))
        valor = float(input("Valor da entrega: "))

        cursor.execute("""
        INSERT INTO entregas (programacao_codigo, cliente_id, caixas, valor)
        VALUES (?, ?, ?, ?)
        """, (codigo_programacao, cliente_id, caixas, valor))

        print("Entrega adicionada com sucesso!")

    conn.commit()
    conn.close()
def adicionar_entregas(codigo_programacao):
    conn = sqlite3.connect("banco.db")
    cursor = conn.cursor()

    while True:
        listar_clientes()
        cliente_id = input("\nDigite o ID do cliente (ou ENTER para finalizar): ")

        if cliente_id == "":
            break

        caixas = int(input("Quantidade de caixas: "))
        valor = float(input("Valor da entrega: "))

        cursor.execute("""
        INSERT INTO entregas (programacao_codigo, cliente_id, caixas, valor)
        VALUES (?, ?, ?, ?)
        """, (codigo_programacao, cliente_id, caixas, valor))

        print("Entrega adicionada com sucesso!")

    conn.commit()
    conn.close()
