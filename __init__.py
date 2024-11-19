import pandas as pd
file_path = "PlanilhaVendas.xlsx"
dados_vendas = pd.read_excel(file_path)

dados_vendas['Receita Total'] = dados_vendas['Quantidade'] * dados_vendas['Preço unitário']

produto_mais_vendido = dados_vendas.groupby('ID do produto')['Quantidade'].sum().idxmax()

receita_por_regiao = dados_vendas.groupby('Região')['Receita Total'].sum()

dia_com_mais_vendas = dados_vendas.groupby('Data da venda')['Quantidade'].sum().idxmax()

print("Produto mais vendido:", produto_mais_vendido)
print("Receita total por região:")
print(receita_por_regiao)
print("Dia com maior número de vendas:", dia_com_mais_vendas)

relatorio_vendas = (
    dados_vendas.groupby('ID do produto')
    .agg(
        Quantidade_Total=('Quantidade', 'sum'),
        Receita_Total=('Receita Total', 'sum')
    )
    .reset_index()
    .sort_values(by='Receita_Total', ascending=False)
)

relatorio_vendas.to_excel("relatorio_vendas.xlsx", index=False)
print("Relatório salvo como 'relatorio_vendas.xlsx'")
