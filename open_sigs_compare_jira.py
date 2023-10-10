import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, Label, Button, Text, Scrollbar
import tkinter.messagebox as messagebox

class BaseApp:
    def centralizar_janela(self, janela, largura, altura):
        largura_tela = janela.winfo_screenwidth()
        altura_tela = janela.winfo_screenheight()
        x = (largura_tela / 2) - (largura / 2)
        y = (altura_tela / 2) - (altura / 2)
        janela.geometry(f'{largura}x{altura}+{int(x)}+{int(y)}')

class TelaIntroducao(BaseApp):
    def __init__(self, mestre):
        self.mestre = mestre
        self.mestre.title("Informações do Processador de Arquivos")
        self.mestre.geometry("700x500")
        self.configurar_interface()

    def configurar_interface(self):
        texto_informativo = """
Delimitadores Sigs = [;,|]
Jira exportar para Better Excel.

Carga dos Arquivos:
- O arquivo export do Sigs é lido csv ou txt.
- O arquivo do Jira é lido em excel.

- Após o processamento, é possível excluir os arquivos export do Sigs e arquivo em aberto da Bare e Holding originais.
"""
        self.centralizar_janela(self.mestre, 700, 500)
        barra_rolagem = Scrollbar(self.mestre)
        barra_rolagem.pack(side=tk.RIGHT, fill=tk.Y)
        txt_info = Text(self.mestre, wrap=tk.WORD, yscrollcommand=barra_rolagem.set)
        txt_info.insert(tk.END, texto_informativo)
        txt_info.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)
        barra_rolagem.config(command=txt_info.yview)
        btn_iniciar = Button(self.mestre, text="Iniciar", command=self.iniciar_aplicativo)
        btn_iniciar.pack(pady=20)

    def iniciar_aplicativo(self):
        self.mestre.destroy()
        raiz = tk.Tk()
        app = Aplicativo(raiz)
        raiz.mainloop()

class Aplicativo(BaseApp):
    def __init__(self, raiz):
        self.raiz = raiz
        self.raiz.title("Comparador Sigs x Jira")
        self.raiz.geometry("400x200")
        self.caminho_csv = None
        self.caminho_xlsx = None
        self.configurar_interface()

    def configurar_interface(self):
        lbl_titulo = Label(self.raiz, text="Bem-vindo ao Comparador Sigs x Jira", font=("Arial Bold", 14))
        lbl_titulo.pack(pady=10)
        btn_processar = Button(self.raiz, text="Processar", command=self.cadastrar_jira)
        btn_processar.pack(pady=10)
        btn_excluir = Button(self.raiz, text="Excluir Arquivos Utilizados", command=self.excluir_arquivos)
        btn_excluir.pack(pady=10)
        self.lbl_status = Label(self.raiz, text="")
        self.lbl_status.pack(pady=10)
        self.centralizar_janela(self.raiz, 400, 200)

    def atualizar_status(self, mensagem):
        self.lbl_status.config(text=mensagem)
        self.raiz.after(3000, self.limpar_status)

    def limpar_status(self):
        self.lbl_status.config(text="")

    def obter_caminho_arquivo(self, titulo, tipo_arquivo, extensao):
        return filedialog.askopenfilename(title=titulo, filetypes=[(tipo_arquivo, extensao)])

    def ler_csv(self, caminho_arquivo):
        linhas_erro = pd.DataFrame(columns=["Linha", "Conteudo"])
        try:
            with open(caminho_arquivo, 'r', encoding='utf-8') as f:
                cabecalho = f.readline().strip()
            possiveis_delimitadores = [";", ",", "|", "\t"]
            delimitador = max(possiveis_delimitadores, key=cabecalho.count)
            num_delimitadores = cabecalho.count(delimitador)
            with open(caminho_arquivo, 'r', encoding='utf-8') as f:
                for i, linha in enumerate(f, start=2):
                    if linha.count(delimitador) != num_delimitadores:  
                        linhas_erro = linhas_erro.append({"Linha": i, "Conteudo": linha.strip()}, ignore_index=True)
            df = pd.read_csv(caminho_arquivo, delimiter=delimitador)
            return df, linhas_erro
        except Exception as e:
            self.atualizar_status(f"Erro ao ler o arquivo: {str(e)}")
            return pd.DataFrame(), linhas_erro

    def ler_excel(self, caminho_arquivo):
        return pd.read_excel(caminho_arquivo)

    def validar_dados_csv(self, df_csv):
        colunas_necessarias = ['Grupo', 'ID']
        for coluna in colunas_necessarias:
            if coluna not in df_csv.columns:
                raise ValueError(f"A coluna '{coluna}' não está presente no arquivo CSV.")
        df_csv['ID'] = df_csv['ID'].str.strip()
        df_csv = df_csv.dropna(subset=['ID'])
        df_csv['ID'] = df_csv['ID'].astype(str)
        df_csv = df_csv[
            (df_csv['Grupo'].str.contains('CAPGEMINI', case=False, na=False)) |
            ((df_csv['Grupo'].str.contains('PEQUENOS ATENDIMENTOS', case=False, na=False)) & 
                (df_csv['Descrição'].str.contains('CAPGEMINI', case=False, na=False)))
        ]
        return df_csv
        
    def validar_dados_excel(self, df_excel):
        colunas_necessarias = ['ID SIGS']
        for coluna in colunas_necessarias:
            if coluna not in df_excel.columns:
                raise ValueError(f"A coluna '{coluna}' não está presente no arquivo Excel.")
        df_excel['ID SIGS'] = df_excel['ID SIGS'].str.strip()
        df_excel = df_excel.dropna(subset=['ID SIGS'])
        df_excel['ID SIGS'] = df_excel['ID SIGS'].astype(str)
        return df_excel
        
    def cadastrar_jira(self):
        self.caminho_csv = self.obter_caminho_arquivo("Selecione o arquivo export do Sigs", "Arquivos CSV/TXT", ("*.csv", "*.txt"))
        if not self.caminho_csv:
            self.atualizar_status("Erro: Arquivo CSV não selecionado.")
            return
        self.caminho_xlsx = self.obter_caminho_arquivo("Selecione o arquivo em aberto da Bare e Holding (Better Excel)", "Arquivos Excel", "*.xlsx")
        if not self.caminho_xlsx:
            self.atualizar_status("Erro: Arquivo Excel não selecionado.")
            return
        try:
            df_csv, linhas_erro_csv = self.ler_csv(self.caminho_csv)
            df_csv = self.validar_dados_csv(df_csv) 
            df_excel = self.ler_excel(self.caminho_xlsx)
            df_excel = self.validar_dados_excel(df_excel) 
            caminho_saida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
            if not caminho_saida:
                self.atualizar_status("Erro: Arquivo de saída não selecionado.")
                return
            with pd.ExcelWriter(caminho_saida, engine='openpyxl') as escritor:
                df_csv.to_excel(escritor, index=False, sheet_name="SIGS")
                df_excel.to_excel(escritor, index=False, sheet_name="JIRA")
                linhas_erro_csv.to_excel(escritor, index=False, sheet_name="ERROS")
                df_para_salvar = df_csv[~df_csv['ID'].isin(df_excel['ID SIGS'])]
                df_para_salvar.to_excel(escritor, index=False, sheet_name="Não Encontrado Jira")
                df_para_fechar = df_excel[~df_excel['ID SIGS'].isin(df_csv['ID'])]
                df_para_fechar.to_excel(escritor, index=False, sheet_name="Não Encontrado Sigs")
            self.atualizar_status("Processo concluído!")
        except FileNotFoundError:
            self.atualizar_status("Erro: Arquivo não encontrado.")
        except pd.errors.EmptyDataError:
            self.atualizar_status("Erro: Arquivo CSV vazio.")
        except pd.errors.ParserError:
            self.atualizar_status("Erro: Problema ao analisar o arquivo CSV.")
        except ValueError as e:
            self.atualizar_status(f"Erro de validação: {str(e)}")
        except Exception as e:
            self.atualizar_status(f"Erro inesperado: {str(e)}")

    def excluir_arquivos(self):
        for caminho_arquivo in [self.caminho_csv, self.caminho_xlsx]:
            if caminho_arquivo and os.path.exists(caminho_arquivo):
                os.remove(caminho_arquivo)
                self.atualizar_status(f"Arquivo {caminho_arquivo} removido com sucesso.")
            else:
                self.atualizar_status(f"Arquivo {caminho_arquivo} não encontrado.")
        self.raiz.after(5000, lambda: self.atualizar_status("Rotina de exclusão finalizada!"))

if __name__ == "__main__":
    raiz_intro = tk.Tk()
    intro = TelaIntroducao(raiz_intro)
    raiz_intro.mainloop()
