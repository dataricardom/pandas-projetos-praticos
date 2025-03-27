import os 
import pandas as pd
from datetime import datetime
def juntar_arquivos_excel_pasta(caminho_pasta, caminho_destino):
    
    # Adicionando timestamp para usar no salvamento do arquivo

    timestamp = datetime.now().strftime('%Y_%m_%d_%H_%M_%S')

    # Lista todos os arquivos de uma pasta

    arquivos = os.listdir(caminho_pasta)

    lista_df = []

    #Carrega cada arquivo excel da pasta
    
    for arquivo in arquivos:
        caminho_arquivo = os.path.join(caminho_pasta, arquivo)
        
        # Carrega o excel para dataframe e adiciona a uma lista de dataframes

        df = pd.read_excel(caminho_arquivo)
        lista_df.append(df)
    #Adiciona os dataframes em um só
    resultado = pd.concat(lista_df, ignore_index=True)
    #Padroniza o nome do arquivo e adiciona uma data.
    nome_arquivo = f'resultado_{timestamp}.xlsx'
    caminho_arquivo_salvo = os.path.join(caminho_destino, nome_arquivo)
    
    #Salva o arquivo na pasta e nome especificados em caminho_arquivo_salvo

    resultado.to_excel(os.path.join(caminho_arquivo_salvo), index=False)

    #Retorna uma mensagem de arquivo salvo e nome do arquivo
    print (f"Incrementação feita com sucesso! Arquivo gerado: {nome_arquivo}")

