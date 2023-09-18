import os
import pandas as pd
import shutil
from datetime import datetime

try:
    # Solicita ao usuário para inserir a data
    data = input('Por favor, insira a data no formato DD/MM/AAAA: ')

    # Converte a string de data para um objeto datetime
    data = datetime.strptime(data, '%d/%m/%Y')

    # Converte o objeto datetime de volta para uma string no formato desejado
    data = data.strftime('%Y%m%d')

    # Abrindo o arquivo Excel
    xls = pd.ExcelFile('caminho_para_arquivo_IT2.xlsx')

    # Lendo as sheets
    df_origem = pd.read_excel(xls, 'origem', header=None)
    df_destino = pd.read_excel(xls, 'destino', header=None)

    # Fechando o arquivo Excel
    xls.close()

    # Obtendo os caminhos das pastas origem e destino
    pastas_origem = df_origem.dropna().values.tolist()
    pasta_destino = df_destino.values[0][0]

    for pasta in pastas_origem:
        encontrou_arquivo = False
        for root, dirs, files in os.walk(pasta[0]):
            for file in files:
                # Extrai a data do nome do arquivo
                data_arquivo = file[-19:-11]
                # Se a data do arquivo corresponder à data inserida pelo usuário
                if data_arquivo == data:
                    encontrou_arquivo = True
                    shutil.copy2(os.path.join(root, file), pasta_destino)
                    print(f'Arquivo {file} copiado')
        if not encontrou_arquivo:
            print(f'Nenhum arquivo encontrado com a data {data} em {pasta[0]}')
except Exception as e:
    print(f'Ocorreu um erro: {e}')

input('Pressione ENTER para sair...')