import pandas as pd
import os

# Defina o caminho relativo para a pasta "planilhas" dentro do projeto
caminho_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_planilhas = os.path.join(caminho_projeto, 'planilhas')

# Nome do arquivo que você deseja ler
nome_arquivo = 'Campari_p4p.xlsx'

# Caminho completo para o arquivo dentro da pasta "planilhas"
caminho_arquivo = os.path.join(caminho_planilhas, nome_arquivo)

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo)
df_sem_primeira_linha = df.drop(index=0)
novos_nomes_colunas = df_sem_primeira_linha.iloc[0]
df_sem_primeira_linha.columns = novos_nomes_colunas
df_final = df_sem_primeira_linha[1:].reset_index(drop=True)

mapeamento_produtos = {
    816: "APEROL",
    9636: "CAMPARI",
    9637: "CAMPARI",
    423: "SAGATIBA",
    683: "SKYY",
    2805: "SKYY",
    8889: "SAGATIBA",
    1209: "SAGATIBA",
    4339: "CAMPARI"
}
mapeamento_equipes = {
    "PILS": "MC",
    "BAR ESPECIAL": "ON",
    "BAGGIO": "MC",
    "KEY ACCOUNT PR OFF": "OFF",
    "TROPICAL": "MC",
    "KEY ACCOUNT PR ON": "ON",
    "CASCAVEL": "MC",
    "MARINGA": "MC",
    "LONDRINA": "MC",
    "EXTRA": "MC",
    "KEY ACCOUNT SC OFF": "OFF_SC",
    "INDUSTRIA E PROMOÇÕES": "MC",
    "SC NORTE": "MC_SC",
    "HEINEKEN": "MC",
    "SC SUL": "MC_SC",
    "TELEVENDAS": "MC",
    "EVENTOS PR": "MC",
    "KEY ACCOUNT SC ON": "ON_SC"
}

tipo_map = {
    35: '5 A 10+ CHECKS',
    36: '5 A 10+ CHECKS'
}

# Função para mapear o tipo de estabelecimento
def map_tipo(estabelecimento):
    return tipo_map.get(estabelecimento, 'OUTROS')

df_final['status'] = df_final['Nome Equipe'].map(mapeamento_equipes)
df_final['MARCA'] = df_final['Produto'].map(mapeamento_produtos)
df_final['Filial'] = df_final['Filial'].astype(int)
df_final['Cliente'] = df_final['Cliente'].astype(int)
df_final['TIPO'] = df_final['Tipo Estabelecimento'].apply(map_tipo)


df_SC = df_final[df_final['Filial'] == 6].reset_index(drop=True)
df_PR = df_final[df_final['Filial'] != 6].reset_index(drop=True)

volumes_PR = df_PR.groupby('MARCA')['Volumes'].sum().reset_index()
volumes_SC = df_SC.groupby('MARCA')['Volumes'].sum().reset_index()

marcas_interesse = ["CAMPARI", "SAGATIBA", "SKYY", "APEROL"]
todas_marcas = pd.DataFrame({'MARCA': marcas_interesse})
volumes_pr = pd.merge(todas_marcas, volumes_PR, on='MARCA', how='left')
volumes_sc = pd.merge(todas_marcas, volumes_SC, on='MARCA', how='left')

volumes_pr['Volumes'] = volumes_pr['Volumes'].fillna(0)
volumes_sc['Volumes'] = volumes_sc['Volumes'].fillna(0)


volumes_pr_sorted = volumes_pr.sort_values(by='MARCA').reset_index(drop=True)
volumes_sc_sorted = volumes_sc.sort_values(by='MARCA').reset_index(drop=True)


Cobertura_PR = df_PR.groupby('TIPO')['Cliente'].nunique().reset_index()
Cobertura_SC = df_SC.groupby('TIPO')['Cliente'].nunique().reset_index()

# ARRUMANDO DF
tipos_esperados = ['5 A 10+ CHECKS', 'OUTROS']
tipos_df = pd.DataFrame({'TIPO': tipos_esperados})
# Garantir que todos os tipos estejam presentes para PR
Cobertura_PR = pd.merge(tipos_df, Cobertura_PR, on='TIPO', how='left').fillna(0)

# Garantir que todos os tipos estejam presentes para SC
Cobertura_SC = pd.merge(tipos_df, Cobertura_SC, on='TIPO', how='left').fillna(0)

# Salvar os DataFrames em um arquivo Excel com planilhas separadas
diretorio = r'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\SCRIPITADOS'

# Nome do arquivo Excel
nome_arquivo = os.path.join(diretorio, 'CAMPARI_P_4_P.xlsx')

# Salvando os DataFrames em um arquivo Excel com múltiplas planilhas
with pd.ExcelWriter(nome_arquivo) as writer:
    volumes_pr_sorted.to_excel(writer, sheet_name='df_volume_PR', index=False)
    volumes_sc_sorted.to_excel(writer, sheet_name='df_volume_SC', index=False)
    Cobertura_PR.to_excel(writer, sheet_name='Cobertura_PR', index=False)
    Cobertura_SC.to_excel(writer, sheet_name='Cobertura_SC', index=False)