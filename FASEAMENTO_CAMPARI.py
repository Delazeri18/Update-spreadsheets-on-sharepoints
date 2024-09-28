import os 
import pandas as pd
import itertools

# Defina o caminho relativo para a pasta "planilhas" dentro do projeto
caminho_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_planilhas = os.path.join(caminho_projeto, 'planilhas')

# Nome do arquivo que você deseja ler
nome_arquivo = 'Campari_equipe.xlsx'

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

tipo_map = {
    35: '5 A 10+ CHECKS',
    36: '5 A 10+ CHECKS'
}

def map_tipo(estabelecimento):
    return tipo_map.get(estabelecimento, 'OUTROS')

df_final['Cliente'] = df_final['Cliente'].astype(int)
df_final['MARCA'] = df_final['Produto'].map(mapeamento_produtos)
df_final['TIPO'] = df_final['Tipo Estabelecimento'].apply(map_tipo)
df_final['Nome Equipe'] = df_final['Nome Equipe'].replace({'BAR ESPECIAL': 'KEY ACCOUNT PR ON'})

# VOLUMES POR MARCA

nome_equipe_unicos_volume = df_final['Nome Equipe'].unique()
marca_unicas_volume = df_final['MARCA'].unique()

# Crie todas as combinações possíveis de 'Nome Equipe' e 'MARCA'
todas_combinacoes_volume = pd.DataFrame(list(itertools.product(nome_equipe_unicos_volume, marca_unicas_volume)),
                                 columns=['Nome Equipe', 'MARCA'])

# Agrupar e somar volumes
result = df_final.groupby(['Nome Equipe', 'MARCA'])['Volumes'].sum().reset_index()

# Combine os resultados do agrupamento com todas as combinações possíveis
result_completo = pd.merge(todas_combinacoes_volume, result, on=['Nome Equipe', 'MARCA'], how='left')

# Preencha valores ausentes com 0
result_completo['Volumes'] = result_completo['Volumes'].fillna(0)

# POSITIVAÇÃO POR MARCA

df_unique_positivacao = df_final.drop_duplicates(subset=['Cliente', 'MARCA', 'Nome Equipe'])

# Obtenha os valores únicos de 'MARCA' e 'Nome Equipe'
marca_unicas_positivacao = df_unique_positivacao['MARCA'].unique()
nome_equipe_unicos_positivacao = df_unique_positivacao['Nome Equipe'].unique()

# Crie todas as combinações possíveis de 'MARCA' e 'Nome Equipe'
todas_combinacoes_positivacao = pd.DataFrame(list(itertools.product(marca_unicas_positivacao, nome_equipe_unicos_positivacao)),
                                 columns=['MARCA', 'Nome Equipe'])

# Agrupar para calcular a quantidade de positivações
positivacoes = df_unique_positivacao.groupby(['MARCA', 'Nome Equipe']).size().reset_index(name='Positivas')

# Combine os resultados do agrupamento com todas as combinações possíveis
positivacoes_completo = pd.merge(todas_combinacoes_positivacao, positivacoes, on=['MARCA', 'Nome Equipe'], how='left')

# Preencha valores ausentes com 0
positivacoes_completo['Positivas'] = positivacoes_completo['Positivas'].fillna(0)

cliente_marcas = df_final.groupby(['Cliente', 'Nome Equipe'])['MARCA'].nunique().reset_index()

# Filtrar os clientes que compraram todas as 4 marcas
clientes_com_4_marcas = cliente_marcas[cliente_marcas['MARCA'] == 4]

# Contar quantos clientes de cada equipe compraram todas as 4 marcas
clientes_por_equipe = clientes_com_4_marcas.groupby('Nome Equipe').size().reset_index(name='Positivas')

# Lista das equipes e marcas desejadas
valores_desejados = [
     'PILS', 'TROPICAL', 'CASCAVEL', 'KEY ACCOUNT PR ON', 'KEY ACCOUNT PR OFF', 'LONDRINA',
    'MARINGA', 'BAGGIO'
]

# Criar um DataFrame com todas as equipes desejadas
todas_equipes = pd.DataFrame(valores_desejados, columns=['Nome Equipe'])

# Combinar o resultado do agrupamento com a lista de equipes desejadas
quarteto = pd.merge(todas_equipes, clientes_por_equipe, on='Nome Equipe', how='left')

# Preencher valores ausentes com 0
quarteto['Positivas'] = quarteto['Positivas'].fillna(0)

# 1. Filtrar os dados para o tipo de estabelecimento "5 A 10+ CHECKS"
estabelecimento_df = df_final[df_final['TIPO'] == '5 A 10+ CHECKS']

# 2. Agrupar por equipe e contar clientes únicos
clientes_por_equipe_estab = estabelecimento_df.groupby('Nome Equipe')['Cliente'].nunique().reset_index()
clientes_por_equipe_estab.columns = ['Nome Equipe', 'Clientes Únicos']

# 3. Lista das equipes desejadas
valores_desejados = [
    'PILS', 'TROPICAL', 'CASCAVEL', 'KEY ACCOUNT PR ON', 'KEY ACCOUNT PR OFF', 'LONDRINA',
    'MARINGA', 'BAGGIO'
]

# 4. Criar um DataFrame com todas as equipes desejadas
todas_equipes = pd.DataFrame(valores_desejados, columns=['Nome Equipe'])

# 5. Combinar o resultado do agrupamento com a lista de equipes desejadas
Positiva_mercado = pd.merge(todas_equipes, clientes_por_equipe_estab, on='Nome Equipe', how='left')

# 6. Preencher valores ausentes com 0
Positiva_mercado['Clientes Únicos'] = Positiva_mercado['Clientes Únicos'].fillna(0)



# Salvar os DataFrames em um arquivo Excel com planilhas separadas
diretorio = r'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\SCRIPITADOS'

# Nome do arquivo Excel
nome_arquivo = os.path.join(diretorio, 'Faseamento_campari.xlsx')

positivacoes_completo = positivacoes_completo.sort_values(by=['Nome Equipe','MARCA']).reset_index(drop=True)
result_completo = result_completo.sort_values(by=['Nome Equipe','MARCA']).reset_index(drop=True)
Positiva_mercado = Positiva_mercado.sort_values(by=['Nome Equipe']).reset_index(drop=True)
quarteto = quarteto.sort_values(by=['Nome Equipe']).reset_index(drop=True)


# Salvando os DataFrames em um arquivo Excel com múltiplas planilhas
with pd.ExcelWriter(nome_arquivo) as writer:
    positivacoes_completo.to_excel(writer, sheet_name='POSITIVA', index=False)
    result_completo.to_excel(writer, sheet_name='VOLUMES', index=False)
    Positiva_mercado.to_excel(writer, sheet_name='5 a 10 check', index=False)
    quarteto.to_excel(writer, sheet_name='QUARTETO', index=False)