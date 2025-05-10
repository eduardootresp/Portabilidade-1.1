from openpyxl import Workbook
from scipy.optimize import fsolve
import os
import pandas as pd
from decimal import Decimal, InvalidOperation

def calcular_prestacao(pv, i, n):
    pv = Decimal(pv)
    i = Decimal(i)
    n = Decimal(n)
    return (pv * i) / (1 - (1 + i) ** -n)

def input_validado(mensagem, tipo=float):
    while True:
        entrada = input(mensagem).strip()

        if tipo in [float, Decimal]:
            entrada = entrada.replace(",", ".")

        try:
            if tipo == Decimal:
                return Decimal(entrada)
            return tipo(entrada)
        except (ValueError, InvalidOperation):
            tipo_nome = "número inteiro" if tipo == int else "número decimal"
            print(f"⚠️ Entrada inválida. Digite um {tipo_nome} válido.")

def cadastrar_emprestimo():
    nome = input("Nome do cliente: ")
    saldo = input_validado("Saldo devedor (R$): ", Decimal)
    meses = input_validado("Parcelas restantes: ", int)
    prestacao_atual = input_validado("Valor da prestação atual (R$): ", Decimal)

    # Resolver taxa de juros com fsolve
    def equacao(i):
        return float(prestacao_atual) - (float(saldo) * i) / (1 - (1 + i)**-meses)

    estimativa_inicial = 0.01
    atual_taxa = fsolve(equacao, estimativa_inicial)[0]
    nova_taxa = input_validado("Nova taxa de juros mensal (%): ", Decimal) / Decimal("100")

    return {
        "nome": nome,
        "saldo_devedor": saldo,
        "parcelas_restantes": meses,
        "prestacao_atual": prestacao_atual,
        "atual_taxa_juros_mensal": float(atual_taxa),
        "nova_taxa_juros_mensal": nova_taxa,
    }

def gerar_planilha_excel(emprestimos):
    wb = Workbook()
    ws = wb.active
    ws.title = "Empréstimos"

    # Dados do contrato atual
    ws.append(['Nome', 'Saldo Devedor', 'Parcelas restantes', 'Taxa de juros atual (%)', 'Prestação atual (R$)'])
    for e in emprestimos:
        ws.append([
            e["nome"],
            e["saldo_devedor"],
            e["parcelas_restantes"],
            round(e["atual_taxa_juros_mensal"] * 100, 2),
            round(e["prestacao_atual"], 2),
        ])

    # Dados com nova taxa
    ws.append([])
    ws.append(["Nome", "Saldo Devedor (R$)", "Parcelas restantes", "Nova taxa de juros (%)", "Nova Prestação (R$)"])
    for e in emprestimos:
        nova_prestacao = calcular_prestacao(e["saldo_devedor"], e["nova_taxa_juros_mensal"], e["parcelas_restantes"])
        ws.append([
            e["nome"],
            e["saldo_devedor"],
            e["parcelas_restantes"],
            round(e["nova_taxa_juros_mensal"] * 100, 2),
            round(nova_prestacao, 2)
        ])

    wb.save("emprestimos.xlsx")
    caminho = os.path.abspath("emprestimos.xlsx")
    print(f"\n📁 Arquivo 'emprestimos.xlsx' gerado com sucesso em:\n{caminho}")

    # Mostrar prévia dos dados da planilha
    try:
        df = pd.read_excel(caminho, sheet_name="Empréstimos")
        df = df.dropna(how='all')  # 👈 remove linhas totalmente vazias
        print("\n📋 Prévia dos dados salvos:")
        print(df.head(10))  # Mostra as 10 primeiras linhas
    except Exception as e:
        print(f"⚠️ Não foi possível mostrar a prévia: {e}")


def main():
    emprestimos = []
    print("=== Sistema de Cálculo de Empréstimos ===")
    print("=== Criado por Eduardo #143217 ===")

    while True:
        print("\nMenu:")
        print("1. Adicionar novo empréstimo")
        print("2. Gerar planilha Excel")
        print("3. Sair")

        escolha = input("Escolha uma opção: ")

        if escolha == "1":
            emprestimo = cadastrar_emprestimo()
            emprestimos.append(emprestimo)

        elif escolha == "2":
            if emprestimos:
                gerar_planilha_excel(emprestimos)
            else:
                print("⚠️ Nenhum empréstimo cadastrado ainda.")

        elif escolha == "3":
            print("Encerrando o programa. Até logo!")
            break

        else:
            print("⚠️ Opção inválida. Tente novamente.")

if __name__ == "__main__":
    main()
