import re
from PyPDF2 import PdfReader
import pandas as pd
import os

def convert_pdf_to_text(pdf_path, txt_path):
    reader = PdfReader(pdf_path)
    with open(txt_path, 'w', encoding='utf-8') as txt_file:
        for page in reader.pages:
            txt_file.write(page.extract_text())

def extract_data_from_text(txt_path):
    with open(txt_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    data = {
        "Custo Bruto": [],
        "Valor Item": [],
        "Valor Unit.": [],
        "Qtde": [],
        "Emb": [],
        "a Receber": [],
        "Seq": [],
        "Pedido de Compras": []
    }

    header_found = False
    pedido_compras = ""
    for line in lines:
        if header_found:
            # Encontrar a expressão "PEDIDO DE COMPRAS"
            if "PEDIDO DE COMPRAS" in line:
                pedido_match = re.search(r'PEDIDO DE COMPRAS\s+(.*?)/L', line)
                if pedido_match:
                    pedido_compras = pedido_match.group(1).strip()
                else:
                    pedido_compras = ""
                continue
            
            # Encontrar a expressão que contenha "/L"
            if "/L" in line:
                break

            # Processar a linha para extrair os campos necessários
            parts = line.strip().split()
            if len(parts) < 7:
                continue  # Ignorar linhas que não têm partes suficientes

            data["Custo Bruto"].append(parts[0])
            data["Valor Item"].append(parts[1])
            data["Valor Unit."].append(parts[2])
            data["Qtde"].append(parts[3])
            
            emb_value = parts[4] + ' ' + parts[5]
            data["Emb"].append(emb_value.strip())
            
            receivable_value = ""
            emb_index = 6
            while emb_index < len(parts) and not parts[emb_index].startswith("UN"):
                receivable_value += ' ' + parts[emb_index]
                emb_index += 1
            data["a Receber"].append(receivable_value.strip())

            seq_value = ""
            while emb_index < len(parts):
                if re.match(r'^\d+$', parts[emb_index]):
                    seq_value = parts[emb_index]
                    break
                emb_index += 1
            data["Seq"].append(seq_value)
            data["Pedido de Compras"].append(pedido_compras)

        if "a Receber Custo Bruto Qtd.Canc. Cod Forn Seq Produtos Emb Qtde Valor Item Valor Unit." in line:
            header_found = True

    return data

def save_to_excel(data, excel_path):
    df = pd.DataFrame(data)
    df.to_excel(excel_path, index=False)

# Caminhos dos arquivos
pdf_path = 'pdf1.pdf'
txt_path = 'output.txt'
excel_path = 'output.xlsx'

# Remover arquivo de saída se existir
if os.path.exists(txt_path):
    os.remove(txt_path)

# Converte PDF para texto
convert_pdf_to_text(pdf_path, txt_path)

# Extrai os dados do arquivo de texto
data = extract_data_from_text(txt_path)

# Salva os dados em um arquivo Excel
save_to_excel(data, excel_path)

print("Arquivo Excel gerado com sucesso!")
