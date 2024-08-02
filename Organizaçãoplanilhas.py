import pandas as pd
import os

def load_and_sort_sheets(folder_path):
    # Lista todos os arquivos no diretório especificado
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    # Lista para armazenar os DataFrames
    data_frames = []
    
    for file in files:
        # Lê cada planilha e adiciona uma coluna com a data do documento
        df = pd.read_excel(os.path.join(folder_path, file))
        df['Document Date'] = pd.to_datetime(df['Document Date'])
        data_frames.append(df)
    
    # Concatena todos os DataFrames
    combined_df = pd.concat(data_frames)
    
    # Ordena pelo campo de data
    combined_df = combined_df.sort_values(by='Document Date')
    
    # Salva o DataFrame combinado e ordenado em um novo arquivo Excel
    combined_df.to_excel('sorted_combined_sheets.xlsx', index=False)

# Exemplo de uso
folder_path = 'coloque\o\diretório\aqui'  # Substitua pelo caminho do seu diretório
load_and_sort_sheets(folder_path)
