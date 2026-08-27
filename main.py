import pandas as pd
import openpyxl
#Processo de dados
df = pd.read_excel("dados/vendas.xlsx")
df["Valor por Unidade"] = df["Valor"]/ df["Quantidade"].round(2)
df["Total"] = df["Total"] + df["Frete"]
total_vendas = df["Quantidade"].sum()
total_produtos = df["Valor"].sum()
total_frete = df["Frete"].sum()
total_geral = df["Total"].sum()
ticket_medio = total_geral / total_vendas
#Criando nova aba para resumo de dados
resumo = pd.DataFrame({
    "Indicador": [
    "Total de vendas",
        "Valor dos produtos",
        "Total de frete",
        "Total geral",
        "Ticket médio"
    ],
    "Valor": [
        total_vendas,
        total_produtos,
        total_frete,
        total_geral,
        ticket_medio
    ]
})
#Saida de excel automatica
with pd.ExcelWriter("saida/relatorio_automatizado.xlsx") as writer:
    df.to_excel(writer, sheet_name="Vendas", index=False)
    resumo.to_excel(writer, sheet_name="Resumo", index=False)