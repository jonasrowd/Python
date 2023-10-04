import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog

root = tk.Tk()
root.withdraw()  

messagebox.showinfo("Instruções", "Por favor, selecione os arquivos de entrada e escolha onde o arquivo de saída será gerado.")

csv_path = filedialog.askopenfilename(title="Selecione o arquivo CSV", filetypes=[("CSV files", "*.csv")])
xlsx_path = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Excel files", "*.xlsx")])

def process_files(csv_path, xlsx_path):

    print("Iniciando o processo...")

    print(f"Lendo o arquivo CSV em {csv_path}...")
    df_csv = pd.read_csv(csv_path, delimiter=';')
    print(f"Registros no CSV: {len(df_csv)}")

    print("Limpando dados do CSV...")
    df_csv['ID'] = df_csv['ID'].str.strip()
    df_csv = df_csv.dropna(subset=['ID'])
    print(f"Registros no CSV após limpeza: {len(df_csv)}")

    print(f"Lendo o arquivo Excel em {xlsx_path}...")
    df_excel = pd.read_excel(xlsx_path)
    print(f"Registros no Excel: {len(df_excel)}")

    print("Limpando dados do Excel...")
    df_excel['ID SIGS'] = df_excel['ID SIGS'].str.strip()
    df_excel = df_excel.dropna(subset=['ID SIGS'])
    print(f"Registros no Excel após limpeza: {len(df_excel)}")

    print("Padronizando tipos de dados das colunas para comparação...")
    df_csv['ID'] = df_csv['ID'].astype(str)
    df_excel['ID SIGS'] = df_excel['ID SIGS'].astype(str)

    print("Comparando dados entre Excel e CSV...")
    ids_nao_encontrados = df_excel['ID SIGS'][~df_excel['ID SIGS'].isin(df_csv['ID'])]
    print(f"IDs não encontrados no CSV: {len(ids_nao_encontrados)}")

    output_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    print(f"Salvando IDs não encontrados em {output_path}...")
    with open(output_path, "w") as file:
        file.write(ids_nao_encontrados.to_string(index=False))

process_files(csv_path, xlsx_path)

delete_files = messagebox.askyesno("Excluir Arquivos", "Deseja excluir os arquivos Excel e CSV após o processamento?")

if delete_files:
    for file_path in [csv_path, xlsx_path]:
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"Arquivo {file_path} removido com sucesso.")
        else:
            print(f"Arquivo {file_path} não encontrado.")

print("Processo concluído!")
