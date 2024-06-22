import os
import time
import P_0_DELETE_TO_REPROCESS_FILES
import P_1_READ_FILE_TO_COLUMN
import P_2_CHOOSE_TO_DIN
import P_2_CHOOSE_TO_FACT
import P_3_EXTRACT_PEDIDO_TO_FACT
import P_3_PROCESS_TO_DIN
import P_4_EXTRACT_N_PRODUCTS_TO_FACT
import P_5_PRODUCT_PEDIDO_TO_FACT
import P_6_FACT

def main():
    try:
        P_0_DELETE_TO_REPROCESS_FILES.run()
    except AttributeError as e:
        print(f"Error in P_0_DELETE_TO_REPROCESS_FILES: {e}")
    time.sleep(5)

    try:
        P_1_READ_FILE_TO_COLUMN.run()
    except AttributeError as e:
        print(f"Error in P_1_READ_FILE_TO_COLUMN: {e}")
    time.sleep(5)

    try:
        P_2_CHOOSE_TO_DIN.run()
    except AttributeError as e:
        print(f"Error in P_2_CHOOSE_TO_DIN: {e}")
    time.sleep(5)

    try:
        P_2_CHOOSE_TO_FACT.run()
    except AttributeError as e:
        print(f"Error in P_2_CHOOSE_TO_FACT: {e}")
    time.sleep(5)

    try:
        P_3_EXTRACT_PEDIDO_TO_FACT.run()
    except AttributeError as e:
        print(f"Error in P_3_EXTRACT_PEDIDO_TO_FACT: {e}")
    time.sleep(5)

    try:
        P_3_PROCESS_TO_DIN.run()
    except AttributeError as e:
        print(f"Error in P_3_PROCESS_TO_DIN: {e}")
    time.sleep(5)

    try:
        P_4_EXTRACT_N_PRODUCTS_TO_FACT.run()
    except AttributeError as e:
        print(f"Error in P_4_EXTRACT_N_PRODUCTS_TO_FACT: {e}")
    time.sleep(5)

    try:
        P_5_PRODUCT_PEDIDO_TO_FACT.run()
    except AttributeError as e:
        print(f"Error in P_5_PRODUCT_PEDIDO_TO_FACT: {e}")
    time.sleep(5)

    try:
        P_6_FACT.run()
    except AttributeError as e:
        print(f"Error in P_6_FACT: {e}")

if __name__ == "__main__":
    main()
