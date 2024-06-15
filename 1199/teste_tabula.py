import tabula
import pandas as pd

# Defina o caminho para o arquivo PDF
pdf_path = r'C:\Users\Glauber Marques\Downloads\Marcilon2\1199\pdf1.pdf'

# Leia a tabela do PDF na primeira página
try:
    # Extraia todas as tabelas na página especificada
    dfs = tabula.read_pdf(pdf_path, pages=1, multiple_tables=True, encoding='ISO-8859-1')
    if dfs:
        df1 = dfs[0]  # Pegue a primeira tabela, se houver mais de uma
        print(df1)
    else:
        print("Nenhuma tabela encontrada na página especificada.")
except Exception as e:
    print(f"Ocorreu um erro ao ler o PDF: {e}")
