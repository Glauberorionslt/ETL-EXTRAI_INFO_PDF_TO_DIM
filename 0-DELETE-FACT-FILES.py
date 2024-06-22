
import os



caminho = [
           r'C:\Users\Glauber Marques\Downloads\Marcilon2\2-CHOOSE_TO_FACT.xlsx',
           r'C:\Users\Glauber Marques\Downloads\Marcilon2\3-EXTRACT_PEDIDO_TO_FACT.xlsx',
           r'C:\Users\Glauber Marques\Downloads\Marcilon2\4-EXTRACT_N_PRODUCTS-TO_FACT.xlsx',
           r"C:\Users\Glauber Marques\Downloads\Marcilon2\5-PRODUCT_PEDIDO_TO_FACT.xlsx",
           r'C:\Users\Glauber Marques\Downloads\Marcilon2\6-FACT.xlsx'
           ]



for files in caminho:
    os.remove(files)