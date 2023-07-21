from openpyxl import workbook, load_workbook

planilha = load_workbook("Produtos.xlsx")
abaAtiva = planilha.active

for celula in abaAtiva ["C"]: 
    if celula.value =="Serviço":
        linha = celula.row
        abaAtiva[f"D3{linha}"] = 1.5
    print(celula.value)
    planilha.save("openpyxl.xlsx")


# abaAtiva["A1"] #Ele separa, coluna inteira "A:A"
