import os
import pdfplumber
import pandas as pd
import INICIAL_SETINGS

# Função para ler o PDF e extrair o texto
def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

# Diretório contendo os arquivos PDF
pdf_directory = INICIAL_SETINGS.diretorio_nativo_dos_arquivos_PDF
# Nome do arquivo Excel de saída
excel_path = os.path.join(INICIAL_SETINGS.pasta_de_processos,"quebra-linha.xlsx")

# Lista para armazenar o texto extraído de todos os PDFs
all_texts = []

# Percorrer todos os arquivos no diretório
for filename in os.listdir(pdf_directory):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(pdf_directory, filename)
        # Extrair texto do PDF
        pdf_text = extract_text_from_pdf(pdf_path)
        # Dividir o texto em linhas e adicionar à lista
        lines = pdf_text.split('\n')
        all_texts.extend(lines)

# Criar um DataFrame do pandas com as linhas como colunas
df = pd.DataFrame([all_texts])

# Salvar o DataFrame no arquivo Excel, substituindo se já existir
df.to_excel(excel_path, index=False, header=False)

print(f"Arquivo Excel '{excel_path}' criado com sucesso!")
