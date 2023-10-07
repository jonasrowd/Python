import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, Label, Button, Text, Scrollbar

class TelaIntroducao:
    def __init__(self, mestre):
        self.mestre = mestre
        self.mestre.title("Informações do Processador de Arquivos")
        self.mestre.geometry("700x500")
        self.configurar_interface()

    def configurar_interface(self):
        texto_informativo = """

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
        self.centralizar_janela(700, 500)

        barra_rolagem = Scrollbar(self.mestre)
        barra_rolagem.pack(side=tk.RIGHT, fill=tk.Y)

        txt_info = Text(self.mestre, wrap=tk.WORD, yscrollcommand=barra_rolagem.set)
        txt_info.insert(tk.END, texto_informativo)
        txt_info.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

        barra_rolagem.config(command=txt_info.yview)

        btn_iniciar = Button(self.mestre, text="Iniciar", command=self.iniciar_aplicativo)
        btn_iniciar.pack(pady=20)

    def centralizar_janela(self, largura, altura):
        largura_tela = self.mestre.winfo_screenwidth()
        altura_tela = self.mestre.winfo_screenheight()
        x = (largura_tela / 2) - (largura / 2)
        y = (altura_tela / 2) - (altura / 2)
        self.mestre.geometry(f'{largura}x{altura}+{int(x)}+{int(y)}')

    def iniciar_aplicativo(self):
        self.mestre.destroy()
        raiz = tk.Tk()
        app = Aplicativo(raiz)
        raiz.mainloop()

class Aplicativo:
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
        self.centralizar_janela(400, 200)

    def centralizar_janela(self, largura, altura):
        largura_tela = self.raiz.winfo_screenwidth()
        altura_tela = self.raiz.winfo_screenheight()
        x = (largura_tela / 2) - (largura / 2)
        y = (altura_tela / 2) - (altura / 2)
        self.raiz.geometry(f'{largura}x{altura}+{int(x)}+{int(y)}')

    def obter_caminho_arquivo(self, titulo, tipo_arquivo, extensao):
        return filedialog.askopenfilename(title=titulo, filetypes=[(tipo_arquivo, extensao)])

    def ler_csv(self, caminho_arquivo):
        return pd.read_csv(caminho_arquivo, delimiter=';')

    def ler_excel(self, caminho_arquivo):
        return pd.read_excel(caminho_arquivo)

    def limpar_dataframe(self, df, nome_coluna):
        df[nome_coluna] = df[nome_coluna].str.strip()
        df = df.dropna(subset=[nome_coluna])
        return df

    def cadastrar_jira(self):
        self.caminho_csv = self.obter_caminho_arquivo("Selecione o arquivo export do Sigs", "Arquivos CSV", "*.csv")
        if not self.caminho_csv:
            self.lbl_status.config(text="Erro: Arquivo CSV não selecionado.")
            return

        self.caminho_xlsx = self.obter_caminho_arquivo("Selecione o arquivo em aberto da Bare e Holding (Better Excel)", "Arquivos Excel", "*.xlsx")
        if not self.caminho_xlsx:
            self.lbl_status.config(text="Erro: Arquivo Excel não selecionado.")
            return

        try:
            df_csv = self.ler_csv(self.caminho_csv)
            df_csv = df_csv[df_csv['Grupo'].str.contains('CAPGEMINI', case=False, na=False)]
            df_csv = self.limpar_dataframe(df_csv, 'ID')
            df_csv['ID'] = df_csv['ID'].astype(str)

            df_excel = self.ler_excel(self.caminho_xlsx)
            df_excel = self.limpar_dataframe(df_excel, 'ID SIGS')
            df_excel['ID SIGS'] = df_excel['ID SIGS'].astype(str)

            caminho_saida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

            if not caminho_saida:
                self.lbl_status.config(text="Erro: Arquivo de saída não selecionado.")
                return

            df_para_salvar = df_csv[~df_csv['ID'].isin(df_excel['ID SIGS'])]
            df_para_salvar.to_excel(caminho_saida, index=False, sheet_name="Cadastrar Jira")

            df_para_fechar = df_excel[~df_excel['ID SIGS'].isin(df_csv['ID'])]
            with pd.ExcelWriter(caminho_saida, engine='openpyxl', mode='a') as escritor:
                df_para_fechar.to_excel(escritor, index=False, sheet_name="Fechar Sigs")

            self.lbl_status.config(text="Processo concluído!")

        except Exception as e:
            self.lbl_status.config(text=f"Erro: {str(e)}")
            return

    def excluir_arquivos(self):
        for caminho_arquivo in [self.caminho_csv, self.caminho_xlsx]:
            if caminho_arquivo and os.path.exists(caminho_arquivo):
                os.remove(caminho_arquivo)
                print(f"Arquivo {caminho_arquivo} removido com sucesso.")
            else:
                print(f"Arquivo {caminho_arquivo} não encontrado.")
        self.lbl_status.config(text="Rotina de exclusáo finalizada!")

if __name__ == "__main__":
    raiz_intro = tk.Tk()
    intro = TelaIntroducao(raiz_intro)
    raiz_intro.mainloop()
