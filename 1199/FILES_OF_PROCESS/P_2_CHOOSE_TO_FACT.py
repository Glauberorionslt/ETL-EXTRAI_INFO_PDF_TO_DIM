import pandas as pd
import os
import re
import time
import INICIAL_SETINGS


# Caminho para o arquivo de entrada e saída
input_file_path = os.path.join(INICIAL_SETINGS.pasta_de_processos,'quebra-linha.xlsx')
output_file_path =os.path.join(INICIAL_SETINGS.pasta_de_processos,'2-CHOOSE_TO_FACT.xlsx')

# Leitura do arquivo Excel
df = pd.read_excel(input_file_path, header=None)

# Encontrar número de colunas entre as expressões
def count_columns_between(df, start_label, end_label):
    header_row = df.iloc[0]
    start_col = None
    end_col = None

    for i, cell in enumerate(header_row):
        if isinstance(cell, str):
            if start_label in cell:
                start_col = i
            elif end_label in cell:
                end_col = i
                break

    if start_col is not None and end_col is not None:
        return end_col - start_col - 1  # Número de colunas entre as expressões
    else:
        return None

# Encontrar o número de colunas entre as expressões
n_products = count_columns_between(df, "Cod Forn Seq Produtos a Receber Emb Qtde Qtd.Canc. Valor Unit. Valor Item Custo Bruto", "DADOS ADICIONAIS ADVERTÊNCIA P/ RECEBIMENTO")

# Inicializar listas para armazenar os dados correspondentes às novas colunas
linha_produtos = []
produtos_a_receber = []
seq = []
qtde = []
qtde_qtd_canc = []
valor_unit = []
valor_item = []
custo_bruto = []
emb = []

# Função para extrair sequência de números
def extract_seq(value):
    match = re.match(r'(\d+)', value)
    return match.group(1) if match else None

# Função para extrair produtos a receber
def extract_produtos(value):
    parts = re.split(r'(\d+)', value, 1)
    if len(parts) > 2:
        produto_part = parts[2].strip()
        match = re.search(r'\bUN\b', produto_part)
        if match:
            produto_a_receber = produto_part[:match.start()].strip()
            return produto_a_receber
    return None

# Função para extrair "Qtde"
def extract_qtde(value):
    if value:
        match = re.search(r'([^,\s]+),\s*([^,\s]*)', value)
        if match:
            qtde_value = match.group(1) + ',' + match.group(2)
            return qtde_value if match.group(2) else None
    return None

# Função para extrair "Valor Unit."
def extract_valor_unit(value, qtde_value):
    if value and qtde_value:
        start_idx = value.find(qtde_value)
        if start_idx != -1:
            remaining_str = value[start_idx + len(qtde_value):].strip()
            valor_unit_match = re.match(r'(\S+)', remaining_str)
            return valor_unit_match.group(1) if valor_unit_match else None
    return None

# Função para extrair "Valor Item."
def extract_valor_item(value, qtde_value, valor_unit_value):
    if value and qtde_value and valor_unit_value:
        start_idx = value.find(qtde_value)
        if start_idx != -1:
            remaining_str = value[start_idx + len(qtde_value):].strip()
            valor_unit_idx = remaining_str.find(valor_unit_value)
            if valor_unit_idx != -1:
                remaining_str = remaining_str[valor_unit_idx + len(valor_unit_value):].strip()
                valor_item_match = re.match(r'(\S+)', remaining_str)
                return valor_item_match.group(1) if valor_item_match else None
    return None

# Função para extrair "Custo Bruto."
def extract_custo_bruto(value, qtde_value, valor_unit_value, valor_item_value):
    if value and qtde_value and valor_unit_value and valor_item_value:
        start_idx = value.find(qtde_value)
        if start_idx != -1:
            remaining_str = value[start_idx + len(qtde_value):].strip()
            valor_unit_idx = remaining_str.find(valor_unit_value)
            if valor_unit_idx != -1:
                remaining_str = remaining_str[valor_unit_idx + len(valor_unit_value):].strip()
                valor_item_idx = remaining_str.find(valor_item_value)
                if valor_item_idx != -1:
                    remaining_str = remaining_str[valor_item_idx + len(valor_item_value):].strip()
                    custo_bruto_match = re.match(r'(\S+)', remaining_str)
                    return custo_bruto_match.group(1) if custo_bruto_match else None
    return None

# Função para extrair "Emb"
def extract_emb(produto_a_receber_value, value, qtde_value):
    if produto_a_receber_value and value and qtde_value:
        start_idx = value.find(produto_a_receber_value)
        if start_idx != -1:
            end_idx = start_idx + len(produto_a_receber_value)
            remaining_str = value[end_idx:].strip()
            qtde_idx = remaining_str.find(qtde_value)
            if qtde_idx != -1:
                return remaining_str[:qtde_idx].strip()
    return None

# Iterar sobre todas as células do DataFrame
for row in df.iterrows():
    for cell in row[1]:
        cell_str = str(cell)
        
        # Verificar se a célula começa com um número
        if cell_str.startswith('8'):
            linha_produtos.append(cell_str)
            seq.append(extract_seq(cell_str))
            produtos_a_receber_value = extract_produtos(cell_str)
            produtos_a_receber.append(produtos_a_receber_value)
            qtde_value = extract_qtde(cell_str)
            qtde.append(qtde_value)
            valor_unit_value = extract_valor_unit(cell_str, qtde_value)
            valor_unit.append(valor_unit_value)
            valor_item_value = extract_valor_item(cell_str, qtde_value, valor_unit_value)
            valor_item.append(valor_item_value)
            custo_bruto_value = extract_custo_bruto(cell_str, qtde_value, valor_unit_value, valor_item_value)
            custo_bruto.append(custo_bruto_value)
            qtde_qtd_canc.append('')  # Adicionar um valor vazio (string vazia)
            emb_value = extract_emb(produtos_a_receber_value, cell_str, qtde_value)
            emb.append(emb_value)

# Certificar que todas as listas têm o mesmo comprimento
max_length = max(len(l) for l in [linha_produtos, seq, produtos_a_receber, qtde, qtde_qtd_canc, valor_unit, valor_item, custo_bruto, emb])

# Função para preencher listas com strings vazias até o tamanho máximo
def pad_list(lst, length):
    return lst + [''] * (length - len(lst))

# Preencher listas com strings vazias até o tamanho máximo
linha_produtos = pad_list(linha_produtos, max_length)
seq = pad_list(seq, max_length)
produtos_a_receber = pad_list(produtos_a_receber, max_length)
qtde = pad_list(qtde, max_length)
qtde_qtd_canc = pad_list(qtde_qtd_canc, max_length)
valor_unit = pad_list(valor_unit, max_length)
valor_item = pad_list(valor_item, max_length)
custo_bruto = pad_list(custo_bruto, max_length)
emb = pad_list(emb, max_length)

# Criar DataFrame resultante
result_df = pd.DataFrame({
    'linha_produtos': linha_produtos,
    'Cod Forn': [''] * max_length,
    'Seq': seq,
    'Produtos a Receber': produtos_a_receber,
    'Qtde': qtde,
    'Qtde Qtd.Canc': qtde_qtd_canc,
    'Valor Unit.': valor_unit,
    'Valor Item': valor_item,
    'Custo Bruto': custo_bruto,
    'Emb': emb,
    'N produtos': [n_products] * max_length  # Adicionar a coluna "N produtos"
})

# Remover o arquivo de saída se já existir
if os.path.exists(output_file_path):
    os.remove(output_file_path)
    time.sleep(3)

# Salvar o DataFrame resultante em um novo arquivo Excel
result_df.to_excel(output_file_path, index=False)

print(f"\nArquivo salvo em {output_file_path}")

# Print valores coletados
print("Valores coletados para 'linha_produtos':", linha_produtos)
print("Valores coletados para 'Seq':", seq)
print("Valores coletados para 'Produtos a Receber':", produtos_a_receber)
print("Valores coletados para 'Qtde':", qtde)

