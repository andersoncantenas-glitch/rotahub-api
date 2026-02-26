import sqlite3

def conferencia_final():
    codigo = input("Digite o código da programação: ")

    print("\n=== CONFERÊNCIA FINAL ===")

    km_saida = int(input("KM de saída: "))
    km_chegada = int(input("KM de chegada: "))

    kg_carregado = float(input("KG carregados: "))
    kg_entregue = float(input("KG entregues: "))

    caixas_entregues = int(input("Caixas entregues: "))
    mortalidade = int(input("Quantidade de mortalidade: "))

    nota_fiscal = input("Número da nota fiscal: ")

    # Cálculo da média de frango (exemplo simples)
    if caixas_entregues > 0:
        media_frango = kg_entregue / caixas_entregues
    else:
        media_frango = 0

    conn = sqlite3.connect("banco.db")
    cursor = conn.cursor()

    cursor.execute("""
    INSERT INTO conferencia
    (programacao_codigo, km_saida, km_chegada, kg_carregado,
     kg_entregue, caixas_entregues, mortalidade, nota_fiscal, media_frango)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        codigo, km_saida, km_chegada, kg_carregado,
        kg_entregue, caixas_entregues, mortalidade, nota_fiscal, media_frango
    ))

    conn.commit()
    conn.close()

    print("\nConferência final registrada com sucesso!")
    print(f"Média de frango: {media_frango:.2f} kg/caixa")
