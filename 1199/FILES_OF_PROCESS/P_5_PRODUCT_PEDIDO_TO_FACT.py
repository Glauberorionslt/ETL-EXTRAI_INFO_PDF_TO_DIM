import pandas as pd
import os
import INICIAL_SETINGS


# Caminhos dos arquivos de entrada e saída
arquivo_pedido_compras = os.path.join(INICIAL_SETINGS.pasta_de_processos,'3-EXTRACT_PEDIDO_TO_FACT.xlsx')
arquivo_resultado_contagem = os.path.join(INICIAL_SETINGS.pasta_de_processos,'4-EXTRACT_N_PRODUCTS-TO_FACT.xlsx')
arquivo_saida = os.path.join(INICIAL_SETINGS.pasta_de_processos,"5-PRODUCT_PEDIDO_TO_FACT.xlsx")

# Ler o arquivo de resultado_contagem para obter as contagens
df_contagem = pd.read_excel(arquivo_resultado_contagem)

# Ler o arquivo de pedido_compras
df_pedidos = pd.read_excel(arquivo_pedido_compras)

# Verificar as colunas disponíveis nos DataFrames
print("Colunas em df_contagem:", df_contagem.columns)
print("Colunas em df_pedidos:", df_pedidos.columns)

# Inicializar listas para armazenar os dados de saída
pedidos_repetidos = []
n_produtos = []

# Iterar sobre as linhas de resultado_contagem e repetir os pedidos conforme N produtos
for idx, row in df_contagem.iterrows():
    pedido = df_pedidos.loc[idx, 'Pedido de compras']
    n = row['N produtos']
    pedidos_repetidos.extend([pedido] * n)
    n_produtos.extend(range(1, n + 1))

# Criar DataFrame de saída
df_saida = pd.DataFrame({
    'Pedido de compras': pedidos_repetidos,
    'N x produtos': n_produtos
})

if os.path.exists(arquivo_saida):
    os.remove(arquivo_saida)


# Salvar o DataFrame de saída em um novo arquivo Excel
df_saida.to_excel(arquivo_saida, index=False)

print(f"Arquivo de saída gerado com sucesso em {arquivo_saida}")
