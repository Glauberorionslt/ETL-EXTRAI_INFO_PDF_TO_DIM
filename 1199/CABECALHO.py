import fitz  # PyMuPDF
import pandas as pd
import re

# Função para extrair texto de cada página do PDF
def extract_text_from_pdf(pdf_file_path):
    pdf_text = ""
    pdf_document = fitz.open(pdf_file_path)
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        pdf_text += page.get_text()
    pdf_document.close()
    return pdf_text

# Função para processar o texto extraído conforme as regras especificadas
def process_text(pdf_text):
    data = {
        "PEDIDO DE COMPRAS": [],
        "FORNECEDOR": [],
        "DADOS PARA FATURAMENTO": [],
        "Endereço": [],
        "Bairro": [],
        "Inscrição Estadual": [],
        "ENDEREÇO PARA COBRANÇA ENDEREÇO PARA ENTREGA": [],
        "Transportador R. Social": [],
        "Endereço Telefone": [],
        "Eans": []
    }
    
    # Utilizamos flags para indicar se estamos processando uma seção específica
    current_section = None
    current_data = []
    
    # Expressões regulares para detectar números e EANs
    number_pattern = re.compile(r'\d+')
    ean_pattern = re.compile(r'EANs:\s*(.*)')
    
    # Dividir o texto por linhas
    lines = pdf_text.splitlines()
    
    for line in lines:
        line = line.strip()
        
        # Verificar se estamos começando uma nova seção
        if any(keyword in line for keyword in data.keys()):
            if current_section is not None:
                # Armazenar dados da seção atual
                data[current_section].append(' '.join(current_data))
                current_data = []
            
            # Encontrar a nova seção
            for keyword in data.keys():
                if keyword in line:
                    current_section = keyword
                    break
        
        # Processar a linha de acordo com a seção atual
        if current_section:
            if current_section == "Eans":
                # Extrair EANs
                match = ean_pattern.search(line)
                if match:
                    ean_value = match.group(1).strip()
                    if ean_value:  # Verificar se não é uma string vazia
                        data[current_section].append(ean_value)
            else:
                # Extrair texto até encontrar uma nova seção ou terminar o documento
                if line and not any(keyword in line for keyword in data.keys()):
                    current_data.append(line)
    
    # Adicionar o último conjunto de dados após o último loop
    if current_section and current_data:
        data[current_section].append(' '.join(current_data))
    
    # Ajustar comprimentos de listas para todas as chaves
    max_length = max(len(data[key]) for key in data.keys())
    for key in data.keys():
        data[key] += [''] * (max_length - len(data[key]))
    
    return data

# Função para criar DataFrame a partir dos dados processados
def create_dataframe(data):
    df = pd.DataFrame(data)
    return df

# Função para salvar DataFrame como arquivo Excel
def save_to_excel(df, excel_file_path):
    df.to_excel(excel_file_path, index=False)
    print(f"Arquivo Excel salvo como {excel_file_path}")

# Caminho para o arquivo PDF
pdf_file_path = 'pdf1.pdf'

# Extrair texto do PDF
pdf_text = extract_text_from_pdf(pdf_file_path)

# Processar o texto extraído
processed_data = process_text(pdf_text)

# Criar DataFrame
df = create_dataframe(processed_data)

# Caminho para o arquivo Excel de saída
excel_file_path = 'trat.xlsx'

# Salvar DataFrame como Excel
save_to_excel(df, excel_file_path)
