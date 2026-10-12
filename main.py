import pandas as pd
#Processo de dados
df = pd.read_excel("dados/vendas.xlsx")
df["Total"] = (df["Valor"] + df["Frete"]).round(2)
df["Valor por Unidade"] = (df["Total"]/ df["Quantidade"]).round(2)
total_vendas = df["Quantidade"].sum().round(2)
total_produtos = df["Valor"].sum().round(2)
total_frete = df["Frete"].sum().round(2)
total_geral = df["Total"].sum().round(2)
ticket_medio = (total_geral / total_vendas).round(2)
#Criando nova aba para resumo de dados
resumo = pd.DataFrame({
    "Indicador": [
    "Total de Vendas",
        "Valor dos Produtos",
        "Total de Frete",
        "Total Geral",
        "Valor Médio por Unidade (c/ Frete)"
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