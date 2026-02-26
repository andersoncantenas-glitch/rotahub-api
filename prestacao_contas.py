import sqlite3

def prestar_contas():
    codigo = input("Digite o código da programação: ")

    conn = sqlite3.connect("banco.db")
    cursor = conn.cursor()

    # Buscar entregas
    cursor.execute("""
    SELECT e.cliente_id, c.nome, e.valor
    FROM entregas e
    JOIN clientes c ON e.cliente_id = c.id
    WHERE e.programacao_codigo = ?
    """, (codigo,))

    entregas = cursor.fetchall()

    if not entregas:
        print("Nenhuma entrega encontrada para esse código.")
        conn.close()
        return

    print("\n=== PAGAMENTOS DOS CLIENTES ===")
    total_recebido = 0

    for cliente_id, nome, valor in entregas:
        print(f"\nCliente: {nome}")
        print(f"Valor da entrega: R$ {valor:.2f}")
        pago = float(input("Valor pago pelo cliente: R$ "))

        cursor.execute("""
        INSERT INTO pagamentos (programacao_codigo, cliente_id, valor_pago)
        VALUES (?, ?, ?)
        """, (codigo, cliente_id, pago))

        total_recebido += pago

    print(f"\nTotal recebido: R$ {total_recebido:.2f}")

    # Despesas
    print("\n=== DESPESAS DA VIAGEM ===")
    total_despesas = 0

    while True:
        descricao = input("Descrição da despesa (ENTER para finalizar): ")

        if descricao == "":
            break

        valor = float(input("Valor da despesa: R$ "))

        cursor.execute("""
        INSERT INTO despesas (programacao_codigo, descricao, valor)
        VALUES (?, ?, ?)
        """, (codigo, descricao, valor))

        total_despesas += valor

    print(f"\nTotal de despesas: R$ {total_despesas:.2f}")

    saldo = total_recebido - total_despesas

    print("\n=== RESULTADO FINAL ===")
    print(f"Total recebido : R$ {total_recebido:.2f}")
    print(f"Total despesas : R$ {total_despesas:.2f}")
    print(f"Saldo em caixa : R$ {saldo:.2f}")

    conn.commit()
    conn.close()
