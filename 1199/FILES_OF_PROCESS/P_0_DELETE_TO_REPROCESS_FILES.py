import os
import INICIAL_SETINGS

caminho = [
    os.path.join(INICIAL_SETINGS.pasta_de_processos,"quebra-linha.xlsx"),
    os.path.join(INICIAL_SETINGS.pasta_de_processos,'2-CHOOSE_TO_FACT.xlsx'),
    os.path.join(INICIAL_SETINGS.pasta_de_processos,'3-EXTRACT_PEDIDO_TO_FACT.xlsx'),
    os.path.join(INICIAL_SETINGS.pasta_de_processos,'4-EXTRACT_N_PRODUCTS-TO_FACT.xlsx'),
    os.path.join(INICIAL_SETINGS.pasta_de_processos,"5-PRODUCT_PEDIDO_TO_FACT.xlsx"),
    os.path.join(INICIAL_SETINGS.pasta_de_processos,'6-FACT.xlsx'),
    os.path.join(INICIAL_SETINGS.pasta_de_processos,'2-CHOOSE_TO_DIN.xlsx'),
    os.path.join(INICIAL_SETINGS.pasta_de_processos,'3-PROCESS_TO_DIN.xlsx'),
    os.path.join(INICIAL_SETINGS.arquivos_finais,'6-FACT.xlsx'),
    os.path.join(INICIAL_SETINGS.arquivos_finais,'3-PROCESS_TO_DIN.xlsx')
   
]

for files in caminho:
    try:
        if os.path.exists(files):
            os.remove(files)
            print(f"Removido: {files}")
        else:
            print(f"Não encontrado: {files}")
    except Exception as e:
        print(f"Erro ao processar {files}: {e}")
