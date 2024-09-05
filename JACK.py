import pandas as pd
import os 

#arquivo = planilha
# Defina o caminho relativo para a pasta "planilhas" dentro do projeto
caminho_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_planilhas = os.path.join(caminho_projeto, 'planilhas')

# Nome do arquivo que você deseja ler
nome_arquivo = 'TEMPLATE_JACK.xlsx'

# Caminho completo para o arquivo dentro da pasta "planilhas"
caminho_arquivo = os.path.join(caminho_planilhas, nome_arquivo)

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo)
df_sem_primeira_linha = df.drop(index=0)
novos_nomes_colunas = df_sem_primeira_linha.iloc[0]
df_sem_primeira_linha.columns = novos_nomes_colunas
df_final = df_sem_primeira_linha[1:].reset_index(drop=True)

df_final['Filial'] = df_final['Filial'].replace({ 6 : 'BIGUAÇU', 7 : 'CURITIBA', 5 : 'CASCAVEL', 3 : 'CAMBE', 4 : 'CURITIBA'})

clientes_distintos_por_filial = df_final.groupby('Filial')['Cliente'].nunique()

clientes_distintos_por_filial = clientes_distintos_por_filial.reset_index()

clientes_distintos_por_filial.columns = ['Filial', 'Clientes Distintos']

Equipes_interesse = ["BIGUAÇU","CASCAVEL","LONDRINA","CAMBE","CURITIBA"]
todas_Equipes = pd.DataFrame({'Filial': Equipes_interesse})

positiva = pd.merge(todas_Equipes, clientes_distintos_por_filial, on='Filial', how='left').fillna(0) # arrumados os df 


positiva_sorted = positiva.sort_values(by='Filial').reset_index(drop=True)



diretorio = r'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\SCRIPITADOS'

nome_arquivo = os.path.join(diretorio, 'JACK_POSITIVAÇÃO.xlsx')
# Nome do arquivo Excel

with pd.ExcelWriter(nome_arquivo) as writer:
    clientes_distintos_por_filial.to_excel(writer, sheet_name='df_volume_PR', index=False)
    