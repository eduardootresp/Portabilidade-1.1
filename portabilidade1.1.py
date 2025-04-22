from openpyxl import Workbook
from scipy.optimize import fsolve
import os

"""Função para calcular o valor da parcela com a nova taxa."""
def calcular_prestacao(pv, i, n):
    return (pv * i) / (1 - (1 + i) ** -n)


'''Função para validar o input do usuário.'''
def input_validado(mensagem, tipo=float):
    while True:
        entrada = input(mensagem).strip()

        if tipo == float:
            entrada = entrada.replace(",", ".")

        try:
            return tipo(entrada)
        except ValueError:
            tipo_nome = "número inteiro" if tipo == int else "número decimal"
            print(f"Entrada inválida. Digite um {tipo_nome} válido.")


'''Função para cadastrar o empréstimo atual.'''
def cadastrar_emprestimo():
    nome = input("Nome do cliente: ")
    saldo = input_validado("Saldo devedor (R$): ", float)
    meses = input_validado("Parcelas restantes: ", int)
    prestacao_atual = input_validado("Valor da prestação atual (R$): ", float)

    # Resolver taxa de juros com fsolve
    def equacao(i):
        return prestacao_atual - (saldo * i) / (1 - (1 + i)**-meses)

    estimativa_inicial = 0.01
    atual_taxa = fsolve(equacao, estimativa_inicial)[0]
    nova_taxa = input_validado("Nova taxa de juros mensal (%): ", float) / 100

    return {
        "nome": nome,
        "saldo_devedor": saldo,
        "parcelas_restantes": meses,
        "prestacao_atual": prestacao_atual,
        "atual_taxa_juros_mensal": float(atual_taxa),
        "nova_taxa_juros_mensal": nova_taxa,
    }


'''Cria as planilhas com os dados do usuário.'''
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

    # mostra onde foi salvo o arquivo
    wb.save("emprestimos.xlsx")
    caminho = os.path.abspath("emprestimos.xlsx")
    print(f"\n Arquivo 'emprestimos.xlsx' gerado com sucesso em:\n{caminho}")

    # Mostrar prévia dos dados da planilha
    from openpyxl import load_workbook

    # ...

    try:
        wb = load_workbook(caminho)
        ws = wb["Empréstimos"]

        print("\n Prévia dos dados salvos:")

        linha_vazia = False
        contador = 0

        for row in ws.iter_rows(values_only=True):
            if all(cell is None for cell in row):
                linha_vazia = True
                continue

            # Formatar e alinhar cada célula
            linha_formatada = " | ".join(
                str(cell).ljust(30) if cell is not None else "".ljust(30)
                for cell in row
            )

            print(linha_formatada)
            contador += 1
            if contador >= 10:
                break

        if linha_vazia:
            print("(Linhas vazias foram ignoradas)")
    except Exception as e:
        print(f"Não foi possível mostrar a prévia: {e}")


'''Faz o menu principal.'''
def main():
    emprestimos = []
    print("=== Sistema de Cálculo de Empréstimos ===")
    print("=== Criado por Eduardo #https://github.com/eduardootresp ===")

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
                print("Nenhum empréstimo cadastrado ainda.")

        elif escolha == "3":
            print("Encerrando o programa. Até logo!")
            break

        else:
            print("Opção inválida. Tente novamente.")

    # Esta linha mantém o terminal aberto até o usuário pressionar Enter
    input("\nPressione Enter para sair...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n Ocorreu um erro inesperado: {e}")
        input("\nPressione Enter para sair...")

