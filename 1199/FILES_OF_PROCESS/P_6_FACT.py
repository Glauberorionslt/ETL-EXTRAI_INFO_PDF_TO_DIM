import pandas as pd
import os
import INICIAL_SETINGS


# Caminhos dos arquivos de entrada e saída
arquivo1 = os.path.join(INICIAL_SETINGS.pasta_de_processos,'2-CHOOSE_TO_FACT.xlsx')
arquivo2 = os.path.join(INICIAL_SETINGS.pasta_de_processos,'5-PRODUCT_PEDIDO_TO_FACT.xlsx')
arquivo_saida = os.path.join(INICIAL_SETINGS.arquivos_finais,'6-FACT.xlsx')

# Ler o arquivo 2-CHOOSE_TO_FACT.xlsx, selecionando as colunas específicas
df1 = pd.read_excel(arquivo1, usecols=['Cod Forn', 'Seq', 'Produtos a Receber', 'Qtde', 'Qtde Qtd.Canc', 
                                       'Valor Unit.', 'Valor Item', 'Custo Bruto', 'Emb'])

# Ler o arquivo saida.xlsx
df2 = pd.read_excel(arquivo2)

# Concatenar os DataFrames ao longo das colunas
df_merged = pd.concat([df1, df2], axis=1)

#verificar existencia arquivo
if os.path.exists(arquivo_saida):
    os.remove(arquivo_saida)


# Salvar o DataFrame concatenado no arquivo de saída
df_merged.to_excel(arquivo_saida, index=False)




print(f"Arquivo de saída gerado com sucesso em {arquivo_saida}.")
