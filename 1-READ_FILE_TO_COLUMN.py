import pdfplumber
import pandas as pd

# Função para ler o PDF e extrair o texto
def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

# Caminho do arquivo PDF e nome do arquivo Excel de saída
pdf_path = "pdf1.pdf"
excel_path = "quebra-linha.xlsx"

# Extrair texto do PDF
pdf_text = extract_text_from_pdf(pdf_path)

# Dividir o texto em linhas
lines = pdf_text.split('\n')

# Criar um DataFrame do pandas com as linhas como colunas
df = pd.DataFrame([lines])

# Salvar o DataFrame no arquivo Excel, substituindo se já existir
df.to_excel(excel_path, index=False, header=False)

print(f"Arquivo Excel '{excel_path}' criado com sucesso!")
