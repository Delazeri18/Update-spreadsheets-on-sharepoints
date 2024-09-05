import pandas as pd
import os

# Defina o caminho relativo para a pasta "planilhas" dentro do projeto
caminho_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_planilhas = os.path.join(caminho_projeto, 'planilhas')

# Nome do arquivo que você deseja ler
nome_arquivo = 'Template_CC.xlsx'

# Caminho completo para o arquivo dentro da pasta "planilhas"
caminho_arquivo = os.path.join(caminho_planilhas, nome_arquivo)

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo)
df_sem_primeira_linha = df.drop(index=0)
novos_nomes_colunas = df_sem_primeira_linha.iloc[0]
df_sem_primeira_linha.columns = novos_nomes_colunas
df_final = df_sem_primeira_linha[1:].reset_index(drop=True)

#mapeando
mapeamento_produtos = {
    816: "APEROL",
    9636: "CAMPARI",
    9637: "CAMPARI",
    423: "SAGATIBA",
    683: "SKYY",
    2805: "SKYY",
    8889: "SAGATIBA",
    1209: "SAGATIBA",
    4339: "CAMPARI",
    8815: "PREMIUM",
    657: "POPULAR CAMPARI",
    6195: "POPULAR CAMPARI",
    2522:"PREMIUM",
    10063: "PREMIUM",
    10064: "PREMIUM",
    10065: "PREMIUM",
    3423: 'PREMIUM',
    539: "PREMIUM",
    925: 'PREMIUM',
    2503 : 'PREMIUM',
    2804 : 'PREMIUM',
    832 : 'PREMIUM',
    2520: 'PREMIUM',
    1522 : 'PREMIUM',
    3426 : 'PREMIUM',
    8733 : 'PREMIUM',
    1399 : 'PREMIUM',
    1398 : 'PREMIUM',
    4327 : 'PREMIUM',
    2523 : "PREMIUM",
    2521 : "PREMIUM",
    716 : 'PREMIUM',
    715 : 'PREMIUM',
    8825 : 'PREMIUM',
    693: 'POPULAR CAMPARI',
    2522 : 'POPULAR CAMPARI',
    388 : 'POPULAR CAMPARI'
}

df_final['MARCA'] = df_final['Produto'].map(mapeamento_produtos)

# filtrando filial

df_SC = df_final[df_final['Filial'] == 6].reset_index(drop=True)
df_PR = df_final[df_final['Filial'] != 6].reset_index(drop=True)

todas_as_marcas = pd.DataFrame({
    'MARCA': ['APEROL', 'CAMPARI', 'SAGATIBA', 'SKYY', 'PREMIUM','POPULAR CAMPARI']  # Inclua todas as marcas possíveis
})

# Cálculo do volume para PR
volumes_PR = df_PR.groupby('MARCA')['Volumes'].sum().reset_index()

# Cálculo do volume para SC
volumes_SC = df_SC.groupby('MARCA')['Volumes'].sum().reset_index()

# Mesclar com todas as marcas para garantir que todas apareçam, mesmo se o volume for zero
volumes_PR = todas_as_marcas.merge(volumes_PR, on='MARCA', how='left').fillna(0)
volumes_SC = todas_as_marcas.merge(volumes_SC, on='MARCA', how='left').fillna(0)

# Arredondar os volumes para 2 casas decimais
volumes_PR['Volumes'] = volumes_PR['Volumes'].round(2)
volumes_SC['Volumes'] = volumes_SC['Volumes'].round(2)

# positivação

positivacao_PR = df_PR.groupby('MARCA')['Cliente'].nunique().reset_index()
positivacao_SC = df_SC.groupby('MARCA')['Cliente'].nunique().reset_index()

# Mesclar com todas as marcas possíveis para garantir que todas apareçam, mesmo se a positivação for zero
positivacao_PR= todas_as_marcas.merge(positivacao_PR, on='MARCA', how='left').fillna(0)
positivacao_SC= todas_as_marcas.merge(positivacao_SC, on='MARCA', how='left').fillna(0)

# Quarteto:
# Lista de todas as marcas que estamos interessados
marcas_necessarias = ['APEROL', 'CAMPARI', 'SAGATIBA', 'SKYY']

# Agrupar por CLIENTE e usar set() para verificar as marcas compradas por cada cliente
QUARTETO_PR = df_PR.groupby('Cliente')['MARCA'].apply(set).reset_index()
QUARTETO_SC = df_SC.groupby('Cliente')['MARCA'].apply(set).reset_index()

# Filtrar clientes que compraram todas as marcas
quarteto_PR = QUARTETO_PR[QUARTETO_PR['MARCA'].apply(lambda x: set(marcas_necessarias).issubset(x))].reset_index()
quarteto_SC = QUARTETO_SC[QUARTETO_SC['MARCA'].apply(lambda x: set(marcas_necessarias).issubset(x))].reset_index()

# Contar o número de clientes que compraram todas as marcas
quarteto_PR1 = quarteto_PR.shape[0]
quarteto_SC1 = quarteto_SC.shape[0]

# volumes totais 
total_volumes_Pr = round(df_PR['Volumes'].sum(),2)
total_volumes_Sc = round(df_SC['Volumes'].sum(),2)

# oranizando df

quarteto_df = pd.DataFrame({
    'Categoria': ['PR', 'SC'],
    'Valor': [quarteto_PR1, quarteto_SC1]
})

volumes = pd.DataFrame({
    'Categoria': ['PR', 'SC'],
    'Volume': [total_volumes_Pr, total_volumes_Sc]
})

volumes_marca = pd.DataFrame({
    'Categoria': ['PR', 'SC'],
    'Volume': [total_volumes_Pr, total_volumes_Sc]
})

# Salvar os DataFrames em um arquivo Excel com planilhas separadas
diretorio = r'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\SCRIPITADOS'

# Nome do arquivo Excel
nome_arquivo = os.path.join(diretorio, 'CAMPARI_CLUB.xlsx')



# Salvando os DataFrames em um arquivo Excel com múltiplas planilhas
with pd.ExcelWriter(nome_arquivo) as writer:
    positivacao_PR.to_excel(writer, sheet_name='df_posi_PR', index=False)
    positivacao_SC.to_excel(writer, sheet_name='df_posi_SC', index=False)
    quarteto_df.to_excel(writer, sheet_name='quarteto', index=False)
    volumes.to_excel(writer, sheet_name='volumes_geral', index=False)
    volumes_SC.to_excel(writer, sheet_name='volumes_marca_SC', index=False)
    volumes_PR.to_excel(writer, sheet_name='volumes_MARCA_PR', index=False)