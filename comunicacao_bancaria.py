import os
from tkinter import filedialog
from tkinter import Tk
import pandas as pd

def extrair_dados(pasta):
    dados = {'retorno': [], 'idcnab': [], 'arquivo': []}

    # Listar todos os arquivos na pasta
    for arquivo in os.listdir(pasta):
        if arquivo.endswith('.ret'):
            caminho_completo = os.path.join(pasta, arquivo)
            
            # Abrir e ler o arquivo
            with open(caminho_completo, 'r') as f:
                for linha in f:
                    # Verificar se as posições 15 e 17 contêm '02', '06' ou '09' e posição 58 a 68 diferente de 0000000000
                    if (linha[15:17] == '02' or linha[15:17] == '06' or linha[15:17] == '09') and linha[58:68] != '0000000000':

                        # Adicionar os dados ao dicionário
                        dados['retorno'].append(linha[15:17])
                        dados['idcnab'].append(linha[58:68])
                        dados['arquivo'].append(arquivo)
                        
    return dados

def salvar_excel(dados, caminho_saida):
    # Converter o dicionário em um DataFrame do pandas
    df = pd.DataFrame(dados)
    
    # Salvar o DataFrame em um arquivo Excel
    df.to_excel(caminho_saida, index=False)

if __name__ == '__main__':
    # Configurar a janela de seleção de pasta para não aparecer
    root = Tk()
    root.withdraw()

    # Pedir para o usuário selecionar a pasta
    pasta = filedialog.askdirectory(title='Selecione a Pasta')

    # Pedir para o usuário selecionar o local e nome do arquivo Excel de saída
    caminho_saida = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')], title='Salvar Arquivo Excel Como')

    # Extrair os dados dos arquivos .ret
    dados = extrair_dados(pasta)

    # Salvar os dados em um arquivo Excel
    salvar_excel(dados, caminho_saida)

    print('Dados extraídos e salvos com sucesso!')
