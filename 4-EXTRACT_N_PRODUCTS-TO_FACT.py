import pandas as pd
import os

# Caminhos do arquivo de entrada e saída
arquivo_entrada = r'C:\Users\Glauber Marques\Downloads\Marcilon2\quebra-linha.xlsx'
arquivo_saida = r'C:\Users\Glauber Marques\Downloads\Marcilon2\4-EXTRACT_N_PRODUCTS-TO_FACT.xlsx'

# Ler o arquivo Excel
df = pd.read_excel(arquivo_entrada, header=None)

# Definir expressões para identificar início e fim da contagem
inicio = "Cod Forn Seq Produtos a Receber Emb Qtde Qtd.Canc. Valor Unit. Valor Item Custo Bruto"
fim = "TOTAIS"

# Inicializar variáveis
contagem_atual = 0
resultados = []

# Iterar pelas células da primeira linha (assumindo que é a única linha)
for celula in df.iloc[0]:
    celula_str = str(celula)
    
    if celula_str == inicio:
        if contagem_atual > 0:
            resultados.append(contagem_atual)
        contagem_atual = 0
        continue
    elif celula_str == fim:
        resultados.append(contagem_atual)
        contagem_atual = 0
        break
    
    if celula_str.startswith("8"):
        contagem_atual += 1

# Adicionar a última contagem
if contagem_atual > 0:
    resultados.append(contagem_atual)

# Criar um novo DataFrame com os resultados das contagens
resultado_df = pd.DataFrame({"N produtos": resultados})

if os.path.exists(arquivo_saida):
    os.remove(arquivo_saida)

# Salvar o resultado em um novo arquivo Excel
resultado_df.to_excel(arquivo_saida, index=False)

print(f"Contagens de colunas iniciadas com '8' entre '{inicio}' e '{fim}' realizadas com sucesso. Resultado salvo em {arquivo_saida}.")
