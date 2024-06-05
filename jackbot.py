import pandas as pd

print("Olá, Jack! \nO que temos hoje?")

caminho_planilha_desorganizada = "./Desorganizado.xlsx"

print("Estou lendo a sua planilha...")

nome_planilha = "Sheet1"

dados_planilha = pd.read_excel(caminho_planilha_desorganizada, sheet_name=nome_planilha)

print("\nTudo certo, a seguir é a sua planilha, vamos organiza-la?\n")
print(dados_planilha)

print("\nVamos verificar se a coluna esta certa:\n")
print(dados_planilha.head())

print("\nTudo certo...aguarde...\n")
dados_organizados = dados_planilha.sort_values(by="Destination Address")

print("\nPlanilha organizada:\n")
print(dados_organizados)

dados_organizados.to_excel("./Organizada.xlsx", index=False)
print("\nHappy happy happy")