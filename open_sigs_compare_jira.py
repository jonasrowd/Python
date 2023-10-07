import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, Label, Button, Text, Scrollbar

class IntroScreen:
    def __init__(self, master):
        self.master = master
        self.master.title("Informações do Processador de Arquivos")
        self.master.geometry("700x500")

        info_text = """

O Arquivo do Sigs deve ser separado por ; e do Jira exportado para Better Excel.

Carrega os Arquivos:

- O arquivo export do Sigs é lido.
- O arquivo arquivo em aberto da Bare e Holding é lido.

Filtragem dos Dados do export do Sigs:

- Filtra as linhas onde a coluna 'Grupo' contém a palavra 'CAPGEMINI'.
- Limpa a coluna 'ID', removendo os espaços e excluindo linhas com valores NaN.
- Converte a coluna 'ID' do export do Sigs para o tipo str.

Limpando os Dados do arquivo em aberto da Bare e Holding:

- Limpa a coluna 'ID SIGS', removendo os espaços e excluindo linhas com valores NaN.
- Converte a coluna 'ID SIGS' para o tipo str.

Comparações:

- IDs do export do Sigs que não estão no arquivo em aberto da Bare e Holding são armazenados na worksheet "Cadastrar Jira".
- IDs do arquivo em aberto da Bare e Holding que não estão no export do Sigs são armazenados na worksheet "Fechar Sigs".

Exclusão de Arquivos:

- Após o processamento, é possível excluir os arquivos export do Sigs e arquivo em aberto da Bare e Holding originais.
"""

        scroll = Scrollbar(self.master)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        txt_info = Text(self.master, wrap=tk.WORD, yscrollcommand=scroll.set)
        txt_info.insert(tk.END, info_text)
        txt_info.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

        scroll.config(command=txt_info.yview)

        btn_start = Button(self.master, text="Iniciar", command=self.start_app)
        btn_start.pack(pady=20)

    def start_app(self):
        self.master.destroy()
        root = tk.Tk()
        app = App(root)
        root.mainloop()
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador Sigs x Jira")
        self.root.geometry("400x200")

        self.csv_path = None
        self.xlsx_path = None

        lbl_title = Label(self.root, text="Bem-vindo ao Comparador Sigs x Jira", font=("Arial Bold", 14))
        lbl_title.pack(pady=10)

        btn_start = Button(self.root, text="Processar", command=self.cadastrar_jira)
        btn_start.pack(pady=10)

        btn_delete = Button(self.root, text="Excluir Arquivos Utilizados", command=self.on_delete_files)
        btn_delete.pack(pady=10)

        self.lbl_status = Label(self.root, text="")
        self.lbl_status.pack(pady=10)

    def get_file_path(self, title, file_type, extension):
        return filedialog.askopenfilename(title=title, filetypes=[(file_type, extension)])

    def read_csv(self, file_path):
        return pd.read_csv(file_path, delimiter=';')

    def read_excel(self, file_path):
        return pd.read_excel(file_path)

    def clean_dataframe(self, df, column_name):
        df[column_name] = df[column_name].str.strip()
        df = df.dropna(subset=[column_name])
        return df

    def save_dataframe(self, df, column_name, compare_column, output_path):
        df[~df[column_name].isin(compare_column)].to_excel(output_path, index=False, sheet_name="Cadastrar Jira")

    def delete_files(self, file_paths):
        for file_path in file_paths:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"Arquivo {file_path} removido com sucesso.")
            else:
                print(f"Arquivo {file_path} não encontrado.")

    def cadastrar_jira(self):
        self.csv_path = self.get_file_path("Selecione o arquivo export do Sigs", "CSV files", "*.csv")
        if not self.csv_path:
            self.lbl_status.config(text="Erro: Arquivo de saída não selecionado.")
            return
        
        self.xlsx_path = self.get_file_path("Selecione o arquivo em aberto da Bare e Holding (Better Excel)", "Excel files", "*.xlsx")
        if not self.xlsx_path:
            self.lbl_status.config(text="Erro: Arquivo de saída não selecionado.")
            return
        
        try:
            df_csv = self.read_csv(self.csv_path)
            df_csv = df_csv[df_csv['Grupo'].str.contains('CAPGEMINI', case=False, na=False)]
            df_csv = self.clean_dataframe(df_csv, 'ID')
            df_csv['ID'] = df_csv['ID'].astype(str)

            df_excel = self.read_excel(self.xlsx_path)
            df_excel = self.clean_dataframe(df_excel, 'ID SIGS')
            df_excel['ID SIGS'] = df_excel['ID SIGS'].astype(str)

            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

            # Verificando se o arquivo de saída foi selecionado
            if not output_path:
                self.lbl_status.config(text="Erro: Arquivo de saída não selecionado.")
                return
        
            # Salvando IDs do CSV não presentes no Excel
            df_to_save = df_csv[~df_csv['ID'].isin(df_excel['ID SIGS'])]
            df_to_save.to_excel(output_path, index=False, sheet_name="Cadastrar Jira")

            # Salvando IDs do Excel não presentes no CSV (Verificação inversa)
            df_to_close = df_excel[~df_excel['ID SIGS'].isin(df_csv['ID'])]
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df_to_close.to_excel(writer, index=False, sheet_name="Fechar Sigs")

            self.lbl_status.config(text="Processo concluído!")

        except Exception as e:
            self.lbl_status.config(text=f"Erro: {str(e)}")
            return

    def on_delete_files(self):
        self.delete_files([self.csv_path, self.xlsx_path])
        self.lbl_status.config(text="Arquivos excluídos com sucesso!")

if __name__ == "__main__":
    root_intro = tk.Tk()
    intro = IntroScreen(root_intro)
    root_intro.mainloop()
