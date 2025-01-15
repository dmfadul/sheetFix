import os


folder_name = "info"
base_filename = "BASE.xlsx"
fonte1_filename = "FONTE 1.xlsx"
fonte2_filename = "FONTE 2.xlsx"


base_tabname = "TESTE P MES 7"
fonte1_tabname = "Sheet1"
fonte2_tabname = "Sheet1"


base_path = os.path.join(os.getcwd(), folder_name, base_filename)
fonte1_filename = os.path.join(os.getcwd(), folder_name, fonte1_filename)
fonte2_filename = os.path.join(os.getcwd(), folder_name, fonte2_filename)


