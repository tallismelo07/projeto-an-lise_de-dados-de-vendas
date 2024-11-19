import pandas as pd

caminho_arquivo = "PlanilhaVendas.xlsx"

dados_vendas = pd.read_excel(caminho_arquivo)

dados_vendas['Receita Total'] = dados_vendas['Quantidade'] * dados_vendas['Preço unitário']

produto_mais_vendido = (
    dados_vendas.groupby('ID do produto')['Quantidade']
    .sum()
    .idxmax()
)

receita_por_regiao = dados_vendas.groupby('Região')['Receita Total'].sum()

dia_mais_vendas = (
    dados_vendas.groupby('Data da venda')['Quantidade']
    .sum()
    .idxmax()
)

print("\n=== Análise de Vendas ===")
print(f"1. Produto mais vendido: {produto_mais_vendido}")
print("\n2. Receita total por região:")
for regiao, receita in receita_por_regiao.items():
    print(f"   - {regiao}: R$ {receita:.2f}")
print(f"\n3. Dia com maior número de vendas: {dia_mais_vendas.strftime('%d/%m/%Y')}")

relatorio_produtos = (
    dados_vendas.groupby('ID do produto')
    .agg(
        Quantidade_Total=('Quantidade', 'sum'),
        Receita_Total=('Receita Total', 'sum')
    )
    .reset_index()
    .sort_values(by='Receita_Total', ascending=False)
)

caminho_relatorio = "relatorio_vendas.xlsx"
relatorio_produtos.to_excel(caminho_relatorio, index=False)

print(f"\nRelatório de vendas salvo com sucesso: {caminho_relatorio}")
