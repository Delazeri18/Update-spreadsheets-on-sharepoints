import pandas as pd
import os

# Defina o caminho relativo para a pasta "planilhas" dentro do projeto
caminho_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_planilhas = os.path.join(caminho_projeto, 'planilhas')

# Nome do arquivo que você deseja ler
nome_arquivo = 'TEMPLATE_BEAM_SUNTORY.xlsx'

# Caminho completo para o arquivo dentro da pasta "planilhas"
caminho_arquivo = os.path.join(caminho_planilhas, nome_arquivo)

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo)

# Remover a primeira linha do DataFrame e redefinir os nomes das colunas
df_sem_primeira_linha = df.drop(index=0)
novos_nomes_colunas = df_sem_primeira_linha.iloc[0]
df_sem_primeira_linha.columns = novos_nomes_colunas
df_final = df_sem_primeira_linha[1:].reset_index(drop=True)

# Substituir valores na coluna 'Filial'
df_final['Filial'] = df_final['Filial'].replace({6: 'BIGUAÇU', 7: 'CENTRO', 5: 'CASCAVEL', 3: 'CAMBE', 4: 'CURITIBA'})

# Contar clientes distintos por filial
clientes_distintos_por_filial = df_final.groupby('Filial')['Cliente'].nunique()



# Salvar os DataFrames em um arquivo Excel com planilhas separadas
diretorio = r'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\SCRIPITADOS'

# Nome do arquivo Excel
nome_arquivo_saida = os.path.join(diretorio, 'BEAM_POSITIVAÇÃO.xlsx')

# Salvar o resultado em um novo arquivo Excel
with pd.ExcelWriter(nome_arquivo_saida) as writer:
    clientes_distintos_por_filial.to_excel(writer, sheet_name='positivações', index=True)
