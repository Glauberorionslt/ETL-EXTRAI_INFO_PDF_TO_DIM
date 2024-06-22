import pandas as pd
import os
import time
import INICIAL_SETINGS

# Definir o caminho do arquivo de entrada e saída
input_file = os.path.join(INICIAL_SETINGS.pasta_de_processos,'quebra-linha.xlsx')
output_file = os.path.join(INICIAL_SETINGS.pasta_de_processos,'3-EXTRACT_PEDIDO_TO_FACT.xlsx')

def buscar_pedidos(input_file):
    # Ler o arquivo Excel
    df = pd.read_excel(input_file, header=None)
    
    # Inicializar uma lista para armazenar os valores de "Pedido de compras"
    pedido_compras = []

    # Iterar sobre todas as células do DataFrame
    for row in df.itertuples(index=False):
        for cell in row:
            if pd.isna(cell):
                continue
            cell_str = str(cell)
            # Coletar dados para "Pedido de compras"
            if 'PEDIDO DE COMPRA' in cell_str.upper():
                pedido_compras.append(cell_str)
    
    return pedido_compras

# Função para garantir que o arquivo de saída seja substituído
def salvar_arquivo(df_resultado, output_file):
    if os.path.exists(output_file):
        os.remove(output_file)
    df_resultado.to_excel(output_file, index=False)
    print(f"\nArquivo salvo em {output_file}")

# Tentar acessar o arquivo até que não esteja mais bloqueado
tentativas = 0
while tentativas < 5:  # Tente até 5 vezes
    try:
        # Buscar valores contendo "PEDIDO DE COMPRA"
        pedidos = buscar_pedidos(input_file)
        
        # Remover valores duplicados
        pedidos = list(set(pedidos))
        
        # Criar um DataFrame com os valores em uma coluna chamada "Pedido de compras"
        df_resultado = pd.DataFrame({'Pedido de compras': pedidos})
        
        # Remover linhas com valores nulos
        df_resultado.dropna(inplace=True)
        
        # Ordenar os valores na coluna "Pedido de compras"
        df_resultado.sort_values(by='Pedido de compras', inplace=True)

        
        
        # Salvar o arquivo de saída, garantindo a substituição se existir
        salvar_arquivo(df_resultado, output_file)
        break

    except PermissionError:
        print(f'Erro: arquivo {input_file} está bloqueado. Tentando novamente...')
        tentativas += 1
        time.sleep(5)  # Aguarda 5 segundos antes de tentar novamente

# Se não foi possível processar após tentativas, exibe mensagem de erro
if tentativas == 5:
    print(f'Erro: Não foi possível acessar o arquivo {input_file} após várias tentativas.')
