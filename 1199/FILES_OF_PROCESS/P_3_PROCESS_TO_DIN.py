import pandas as pd
import numpy as np
import os
import time
import INICIAL_SETINGS

# Definir o caminho do arquivo de entrada e saída
input_file = os.path.join(INICIAL_SETINGS.pasta_de_processos,"2-CHOOSE_TO_DIN.xlsx")
output_file = os.path.join(INICIAL_SETINGS.arquivos_finais,"3-PROCESS_TO_DIN.xlsx")

# Tentar acessar o arquivo até que não esteja mais bloqueado
tentativas = 0
while tentativas < 5:  # Tente até 5 vezes
    try:
        # Ler o arquivo Excel, carregando apenas a coluna "Pedido de compras" e excluir valores repetidos
        df_pedido_compras = pd.read_excel(input_file, usecols=['Pedido de compras'])
        df_pedido_compras = df_pedido_compras.drop_duplicates().reset_index(drop=True)

        # Remover a última linha se estiver vazia
        if pd.isnull(df_pedido_compras.iloc[-1]['Pedido de compras']):
            df_pedido_compras = df_pedido_compras[:-1]

        # Ler o arquivo Excel, carregando apenas a coluna "Fornecedor" e manter a quantidade de valores após tratamento
        df_fornecedor = pd.read_excel(input_file, usecols=['Fornecedor'])

        # Excluir valores duplicados na coluna "Fornecedor" e pegar o valor único
        df_fornecedor_unique = df_fornecedor['Fornecedor'].drop_duplicates().reset_index(drop=True).iloc[0]

        # Replicar o valor único na mesma quantidade de linhas de "Pedido de compras"
        num_rows = len(df_pedido_compras)
        df_fornecedor_replicated = pd.DataFrame({'Fornecedor': np.repeat(df_fornecedor_unique, num_rows)})

        # Ler o arquivo Excel, carregando apenas a coluna "Nome_forn" e excluir valores duplicados, aplicando regra específica
        df_nome_forn = pd.read_excel(input_file, usecols=['Nome_forn'])

        # Remover valores duplicados na coluna "Nome_forn" e filtrar linhas vazias e que contenham apenas "R. Social"
        df_nome_forn = df_nome_forn.drop_duplicates(subset=['Nome_forn']).reset_index(drop=True)
        df_nome_forn = df_nome_forn[df_nome_forn['Nome_forn'].astype(str).str.strip() != 'R. Social']
        df_nome_forn = df_nome_forn.dropna().reset_index(drop=True)

        # Ler o arquivo Excel, carregando apenas a coluna "Endereço_fat" e aplicar regras específicas
        df_endereco_fat = pd.read_excel(input_file, usecols=['Endereço_fat'])

        # Excluir valores duplicados na coluna "Endereço_fat" e filtrar linhas que contêm apenas "Endereço Telefone"
        df_endereco_fat = df_endereco_fat.drop_duplicates().reset_index(drop=True)
        df_endereco_fat = df_endereco_fat[~((df_endereco_fat['Endereço_fat'].str.strip() == 'Endereço Telefone') & (df_endereco_fat['Endereço_fat'].str.len() == len('Endereço Telefone')))]
        df_endereco_fat = df_endereco_fat.dropna().reset_index(drop=True)

        # Ler o arquivo Excel, carregando apenas a coluna "cidade_fat" e aplicar regras específicas
        df_cidade_fat = pd.read_excel(input_file, usecols=['cidade_fat'])

        # Excluir valores duplicados na coluna "cidade_fat", filtrar linhas vazias, remover linhas que começam com 4 números
        df_cidade_fat = df_cidade_fat.drop_duplicates().reset_index(drop=True)
        df_cidade_fat = df_cidade_fat.dropna().reset_index(drop=True)
        df_cidade_fat = df_cidade_fat[~df_cidade_fat['cidade_fat'].astype(str).str.match(r'^\d{4}')]

        # Remover linhas vazias da coluna "cidade_fat"
        df_cidade_fat = df_cidade_fat.dropna().reset_index(drop=True)

        # Ler o arquivo Excel, carregando apenas a coluna "cep_fat" e aplicar regras específicas
        df_cep_fat = pd.read_excel(input_file, usecols=['cep_fat'])

        # Excluir valores duplicados na coluna "cep_fat", filtrar linhas vazias, remover linhas que começam com 4 números
        df_cep_fat = df_cep_fat.drop_duplicates().reset_index(drop=True)
        df_cep_fat = df_cep_fat.dropna().reset_index(drop=True)
        df_cep_fat = df_cep_fat[~df_cep_fat['cep_fat'].astype(str).str.match(r'^\d{4}')]

        # Remover linhas que contêm "EANs:" na coluna "cep_fat"
        df_cep_fat = df_cep_fat[~df_cep_fat['cep_fat'].astype(str).str.contains('EANs:', na=False)]

        # Remover linhas vazias da coluna "cep_fat"
        df_cep_fat = df_cep_fat.dropna().reset_index(drop=True)

        # Ler o arquivo Excel, carregando apenas a coluna "end_entr/cobr" e aplicar regras específicas
        df_end_entr_cobr = pd.read_excel(input_file, usecols=['end_entr/cobr'])

        # Excluir linhas vazias da coluna "end_entr/cobr"
        df_end_entr_cobr = df_end_entr_cobr.dropna().reset_index(drop=True)

        # Ler o arquivo Excel, carregando apenas a coluna "Bairro_entr/cobr" e aplicar regras específicas
        df_bairro_entr_cobr = pd.read_excel(input_file, usecols=['Bairro_entr/cobr'])

        # Excluir linhas vazias da coluna "Bairro_entr/cobr"
        df_bairro_entr_cobr = df_bairro_entr_cobr.dropna().reset_index(drop=True)

        # Ler o arquivo Excel, carregando apenas a coluna "Cidade_entr/cobr" e aplicar regras específicas
        df_cidade_entr_cobr = pd.read_excel(input_file, usecols=['Cidade_entr/cobr'])

        # Excluir linhas vazias da coluna "Cidade_entr/cobr"
        df_cidade_entr_cobr = df_cidade_entr_cobr.dropna().reset_index(drop=True)

        # Ler o arquivo Excel, carregando apenas a coluna "Transportador" e aplicar regras específicas
        df_transportador = pd.read_excel(input_file, usecols=['Transportador'])

        # Excluir linhas vazias da coluna "Transportador"
        df_transportador = df_transportador.dropna().reset_index(drop=True)

        # Ler o arquivo Excel, carregando apenas a coluna "R. Social_transp" e aplicar regras específicas
        df_rsocial_transp = pd.read_excel(input_file, usecols=['R. Social_transp'])

        # Excluir linhas vazias da coluna "R. Social_transp"
        df_rsocial_transp = df_rsocial_transp.dropna().reset_index(drop=True)

        # Ler o arquivo Excel, carregando apenas a coluna "Endereço Telefone" e aplicar regras específicas
        df_endereco_telefone = pd.read_excel(input_file, usecols=['Endereço Telefone'])

        # Excluir linhas vazias da coluna "Endereço Telefone"
        df_endereco_telefone = df_endereco_telefone.dropna().reset_index(drop=True)

        # Combinar as colunas tratadas em um único DataFrame
        df_final = pd.concat([df_pedido_compras, df_fornecedor_replicated, df_nome_forn, df_endereco_fat, df_cidade_fat, df_cep_fat, df_end_entr_cobr, df_bairro_entr_cobr, df_cidade_entr_cobr, df_transportador, df_rsocial_transp, df_endereco_telefone], axis=1)

        # Remover todas as linhas sobressalentes onde "Pedido de compras" está vazio
        df_final = df_final.dropna(subset=['Pedido de compras']).reset_index(drop=True)

        # Verificar se o arquivo de saída já existe e, se existir, remover
        if os.path.exists(output_file):
            os.remove(output_file)

        # Salvar o arquivo de saída
        df_final.to_excel(output_file, index=False)

        print(f"\nArquivo salvo em {output_file}")
        break

    except PermissionError:
        print(f'Erro: arquivo {input_file} está bloqueado. Tentando novamente...')
        tentativas += 1
        time.sleep(5)  # Aguarda 5 segundos antes de tentar novamente

    except ValueError as e:
        print(f'Erro ao carregar coluna: {e}')
        break

# Se não foi possível processar após tentativas, exibe mensagem de erro
if tentativas == 5:
    print(f'Erro: Não foi possível acessar o arquivo {input_file} após várias tentativas.')
