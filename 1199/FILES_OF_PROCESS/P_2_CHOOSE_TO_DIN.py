import pandas as pd
import os
import INICIAL_SETINGS

# Caminho para o arquivo de entrada e saída
input_file_path = os.path.join(INICIAL_SETINGS.pasta_de_processos,"quebra-linha.xlsx")
output_file_path = os.path.join(INICIAL_SETINGS.pasta_de_processos,'2-CHOOSE_TO_DIN.xlsx')

# Leitura do arquivo Excel
df = pd.read_excel(input_file_path, header=None)

# Inicializar um novo DataFrame para armazenar as colunas processadas
result_df = pd.DataFrame()

# Inicializar listas para armazenar os dados correspondentes às novas colunas
pedido_compras = []
fornecedor = []
nome_forn = []
endereco_fat = []
bairro_fat = []
cidade_fat = []
cep_fat = []
end_entrega_titulo = []
end_entr_cobr = []
bairro_entr_cobr = []
cidade_entr_cobr = []
transportador = []
rsocial_transp = []
endereco_telefone = []

# Função para adicionar valor às listas correspondentes
def add_to_list(value, lists, index):
    if index < len(lists):
        lists[index].append(value)
    else:
        for lst in lists:
            lst.append(None)
        lists[index].append(value)

# Iterar sobre todas as células do DataFrame
for row in df.itertuples(index=False):
    for i, cell in enumerate(row):
        if pd.isna(cell):
            continue
        cell_str = str(cell)

        # 1. Coletar dados para "Pedido de compras"
        if cell_str.startswith('PEDIDO DE COMPRA'):
            add_to_list(cell_str, [pedido_compras], 0)

        # 2. Coletar dados para "Fornecedor"
        elif cell_str.startswith('FORNECEDOR'):
            add_to_list(cell_str, [fornecedor], 0)

        # 3. Coletar dados para "Nome_forn"
        elif cell_str.startswith('R. Social'):
            add_to_list(cell_str, [nome_forn], 0)

            # Usar i para pegar as células subsequentes
            if i + 1 < len(row):
                endereco_value = row[i + 1]
                add_to_list(endereco_value, [endereco_fat], 0)

            if i + 2 < len(row):
                bairro_value = row[i + 2]
                add_to_list(bairro_value, [bairro_fat], 0)

            if i + 3 < len(row):
                cidade_value = row[i + 3]
                add_to_list(cidade_value, [cidade_fat], 0)

            if i + 4 < len(row):
                cep_value = row[i + 4]
                add_to_list(cep_value, [cep_fat], 0)

        # 4. Coletar dados para "end_entrega_titulo"
        elif cell_str.startswith('ENDEREÇO PARA ENTREGA ENDEREÇO PARA COBRANÇA'):
            add_to_list(cell_str, [end_entrega_titulo], 0)

            # 5. Coletar dados para "end_entr/cobr" (próxima célula após "ENDEREÇO PARA ENTREGA ENDEREÇO PARA COBRANÇA")
            if i + 1 < len(row):
                end_entr_cobr_value = row[i + 1]
                add_to_list(end_entr_cobr_value, [end_entr_cobr], 0)

            # 6. Coletar dados para "Bairro_entr/cobr" (segunda célula após "ENDEREÇO PARA ENTREGA ENDEREÇO PARA COBRANÇA")
            if i + 2 < len(row):
                bairro_entr_cobr_value = row[i + 2]
                add_to_list(bairro_entr_cobr_value, [bairro_entr_cobr], 0)

            # 7. Coletar dados para "Cidade_entr/cobr" (terceira célula após "ENDEREÇO PARA ENTREGA ENDEREÇO PARA COBRANÇA")
            if i + 3 < len(row):
                cidade_entr_cobr_value = row[i + 3]
                add_to_list(cidade_entr_cobr_value, [cidade_entr_cobr], 0)

        # 8. Coletar dados para "Transportador"
        elif 'Transportador' in cell_str:
            add_to_list(cell_str, [transportador], 0)

            # 9. Coletar dados para "R. Social_transp" (próxima célula após "Transportador")
            if i + 1 < len(row):
                rsocial_transp_value = row[i + 1]
                add_to_list(rsocial_transp_value, [rsocial_transp], 0)

        # 10. Coletar dados para "Endereço Telefone"
        elif 'Endereço Telefone' in cell_str:
            add_to_list(cell_str, [endereco_telefone], 0)

# Certificar que todas as listas têm o mesmo comprimento
max_length = max(len(l) for l in [pedido_compras, fornecedor, nome_forn, endereco_fat, bairro_fat, cidade_fat, cep_fat, end_entrega_titulo, end_entr_cobr, bairro_entr_cobr, cidade_entr_cobr, transportador, rsocial_transp, endereco_telefone])

def pad_list(lst, length):
    return lst + [None] * (length - len(lst))

pedido_compras = pad_list(pedido_compras, max_length)
fornecedor = pad_list(fornecedor, max_length)
nome_forn = pad_list(nome_forn, max_length)
endereco_fat = pad_list(endereco_fat, max_length)
bairro_fat = pad_list(bairro_fat, max_length)
cidade_fat = pad_list(cidade_fat, max_length)
cep_fat = pad_list(cep_fat, max_length)
end_entrega_titulo = pad_list(end_entrega_titulo, max_length)
end_entr_cobr = pad_list(end_entr_cobr, max_length)
bairro_entr_cobr = pad_list(bairro_entr_cobr, max_length)
cidade_entr_cobr = pad_list(cidade_entr_cobr, max_length)
transportador = pad_list(transportador, max_length)
rsocial_transp = pad_list(rsocial_transp, max_length)
endereco_telefone = pad_list(endereco_telefone, max_length)

# Adicionar os dados coletados ao DataFrame resultante
result_df['Pedido de compras'] = pedido_compras
result_df['Fornecedor'] = fornecedor
result_df['Nome_forn'] = nome_forn
result_df['Endereço_fat'] = endereco_fat
result_df['Bairro_fat'] = bairro_fat
result_df['cidade_fat'] = cidade_fat
result_df['cep_fat'] = cep_fat
result_df['end_entrega_titulo'] = end_entrega_titulo
result_df['end_entr/cobr'] = end_entr_cobr
result_df['Bairro_entr/cobr'] = bairro_entr_cobr
result_df['Cidade_entr/cobr'] = cidade_entr_cobr
result_df['Transportador'] = transportador
result_df['R. Social_transp'] = rsocial_transp
result_df['Endereço Telefone'] = endereco_telefone

# Remover o arquivo de saída se já existir
if os.path.exists(output_file_path):
    os.remove(output_file_path)

# Salvar o DataFrame resultante em um novo arquivo Excel
result_df.to_excel(output_file_path, index=False)

print(f"\nArquivo salvo em {output_file_path}")
