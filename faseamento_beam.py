import os 
import pandas as pd

# Defina o caminho relativo para a pasta "planilhas" dentro do projeto
caminho_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_planilhas = os.path.join(caminho_projeto, 'planilhas')

# Nome do arquivo que você deseja ler
nome_arquivo = 'Equipe_Beam.xlsx'

# Caminho completo para o arquivo dentro da pasta "planilhas"
caminho_arquivo = os.path.join(caminho_planilhas, nome_arquivo)

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo)

df_sem_primeira_linha = df.drop(index=0)
novos_nomes_colunas = df_sem_primeira_linha.iloc[0]
df_sem_primeira_linha.columns = novos_nomes_colunas
df_final = df_sem_primeira_linha[1:].reset_index(drop=True)

df_final['Cliente'] = df_final['Cliente'].astype(int)
df_final['Nome Equipe'] = df_final['Nome Equipe'].replace({'BAR ESPECIAL': 'KEY ACCOUNT PR ON'})

positivacao = df_final.groupby('Nome Equipe')['Cliente'].nunique().reset_index()
volume = df_final.groupby('Nome Equipe')['Volumes'].sum().reset_index()

positivacao.columns = ['Nome Equipe', 'Clientes Distintos']
volume.columns = ['Nome Equipe', 'Volumes']

volume = volume.reset_index(drop=True)
positivacao = positivacao.reset_index(drop=True)

Equipes_interesse = ["PILS", "BAGGIO", "TROPICAL", "CASCAVEL","MARINGA","LONDRINA","KEY ACCOUNT PR ON","KEY ACCOUNT PR OFF"]
todas_Equipes = pd.DataFrame({'Nome Equipe': Equipes_interesse})
todas_volumes = pd.DataFrame({'Nome Equipe': Equipes_interesse})


positiva = pd.merge(todas_Equipes, positivacao, on='Nome Equipe', how='left').fillna(0)
volumes = pd.merge(todas_volumes, volume, on="Nome Equipe", how='left').fillna(0)



volumes_sorted = volumes.sort_values(by='Nome Equipe').reset_index(drop=True)
positiva_sorted = positiva.sort_values(by='Nome Equipe').reset_index(drop=True)


# Salvar os DataFrames em um arquivo Excel com planilhas separadas
diretorio = r'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\SCRIPITADOS'

# Nome do arquivo Excel
nome_arquivo = os.path.join(diretorio, 'faseamento_beam.xlsx')

# Salvando os DataFrames em um arquivo Excel com múltiplas planilhas
with pd.ExcelWriter(nome_arquivo) as writer:
    volumes_sorted.to_excel(writer, sheet_name='df_volume', index=False)
    positiva_sorted.to_excel(writer, sheet_name='positivacao', index=False)




