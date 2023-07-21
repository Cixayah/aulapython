import pandas as pd

tabela = pd.read_excel("Produtos.xlsx")
tabela.loc[tabela["Tipo"] == "Serviço", "Multiplicador Imposto"] = 1.5
tabela["Preço Base Reais"] = (
    tabela["Multiplicador Imposto"] * tabela["Preço Base Original"]
)
print(tabela)
tabela.to_excel("Pandas.xlsx", index=False)
