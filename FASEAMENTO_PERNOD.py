import pandas as pd
import itertools
import os

# Defina o caminho relativo para a pasta "planilhas" dentro do projeto
caminho_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_planilhas = os.path.join(caminho_projeto, 'planilhas')

# Nome do arquivo que você deseja ler
nome_arquivo = 'PERNOD_EQUIPES.xlsx'

# Caminho completo para o arquivo dentro da pasta "planilhas"
caminho_arquivo = os.path.join(caminho_planilhas, nome_arquivo)

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo)
df_sem_primeira_linha = df.drop(index=0)
novos_nomes_colunas = df_sem_primeira_linha.iloc[0]
df_sem_primeira_linha.columns = novos_nomes_colunas
df_final = df_sem_primeira_linha[1:].reset_index(drop=True)

# REALAIZANDO MAPEAMENTO

mapeamento = {
    444: "SAO FRANCISCO",
    690: "DOMECQ",
    8856: "BEEFEATER",
    689: "ABSOLUT",
    6883: "ABSOLUT",
    3726: "ABSOLUT",
    1560: "ABSOLUT",
    1348: "ABSOLUT",
    684: "ORLOFF",
    723: "BALLANTINES",
    1561: "BALLANTINES",
    4957: "BALLANTINES",
    740: "CHIVAS",
    1562: "CHIVAS",
    4265: "NATU",
    4276: "NATU",
    743: "PASSPORT",
    4738: "PASSPORT",
    8929: "BEEFEATER",
    8928: "BEEFEATER",
    4574: "BEEFEATER",
    9791: "BEEFEATER",
    1298: "BEEFEATER",
    2347: "BEEFEATER",
    4145: "ORLOFF",
    4358: "BALLANTINES",
    1524: "ABSOLUT",
    3569: "ABSOLUT",
    4305: "ABSOLUT",
    1317: "ABSOLUT",
    1508: "ABSOLUT",
    2269: "ABSOLUT",
    8888: "ABSOLUT",
    6798: "ABSOLUT",
    6814: "ABSOLUT",
    1525: "ABSOLUT",
    2948: "ABSOLUT",
    5047: "ORLOFF",
    708: "ORLOFF",
    9725: "BALLANTINES",
    9770: "BALLANTINES",
    709: "BALLANTINES",
    1507: "BALLANTINES",
    1232: "BALLANTINES",
    4273: "BALLANTINES",
    9567: "BALLANTINES",
    84103: "BALLANTINES",
    9566: "BALLANTINES",
    9707: "CHIVAS",
    1746: "CHIVAS",
    4274: "CHIVAS",
    1526: "CHIVAS",
    3510: "CHIVAS"
}


mapeamento_categoria = {
    "SAO FRANCISCO": "NACIONAL",
    "DOMECQ": "NACIONAL",
    "PASSPORT": "NACIONAL",
    "ORLOFF": "NACIONAL",
    "NATU": "NACIONAL"
}

# Função para aplicar a lógica
def classificar_categoria(marca):
    return mapeamento_categoria.get(marca, "OUTRO")

# ADICIONANDO CRITÉRIOS

df_final['Cliente'] = df_final['Cliente'].astype(int)
df_final['MARCA'] = df_final['Produto'].map(mapeamento)
df_final['ANALISE'] = df_final['MARCA'].apply(classificar_categoria)
df_final['Nome Equipe'] = df_final['Nome Equipe'].replace({'BAR ESPECIAL': 'KEY ACCOUNT PR ON'})

nome_equipe_unicos = df_final['Nome Equipe'].unique()
marca_unicas = df_final['MARCA'].unique()

# Criar todas as combinações possíveis de 'Nome Equipe' e 'MARCA'
todas_combinacoes = pd.DataFrame(
    list(itertools.product(nome_equipe_unicos, marca_unicas)),
    columns=['Nome Equipe', 'MARCA']
)

# Agrupar e somar volumes
result = df_final.groupby(['Nome Equipe', 'MARCA'])['Volumes'].sum().reset_index()

# Combinar os resultados do agrupamento com todas as combinações possíveis
result_completo = pd.merge(todas_combinacoes, result, on=['Nome Equipe', 'MARCA'], how='left')

# Preencher valores ausentes com 0
result_completo['Volumes'] = result_completo['Volumes'].fillna(0)

nome_equipe_unicos_analise = df_final['Nome Equipe'].unique()
marca_unicas_analise = df_final['ANALISE'].unique()

# Criar todas as combinações possíveis de 'Nome Equipe' e 'MARCA'
todas_combinacoes_analise = pd.DataFrame(
    list(itertools.product(nome_equipe_unicos_analise, marca_unicas_analise)),
    columns=['Nome Equipe', 'ANALISE']
)

# Agrupar e somar volumes
result_nacional = df_final.groupby(['Nome Equipe', 'ANALISE'])['Volumes'].sum().reset_index()

# Combinar os resultados do agrupamento com todas as combinações possíveis
result_completo_nacional = pd.merge(todas_combinacoes_analise, result_nacional, on=['Nome Equipe', 'ANALISE'], how='left')

# Preencher valores ausentes com 0
result_completo_nacional['Volumes'] = result_completo_nacional['Volumes'].fillna(0)


# volume total por equipe 
total_equipes = df_final.groupby('Nome Equipe')['Volumes'].sum().reset_index()

# Renomeando as colunas corretamente
total_equipes.columns = ['Nome Equipe', 'Volumes']

# Resetando o índice e removendo o índice anterior
total_equipes = total_equipes.reset_index(drop=True)



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

# positivação marca inteira

pernod_posi = positivacoes = df_unique_positivacao.groupby('Nome Equipe')['Cliente'].nunique()

pernod_posi = pernod_posi.reset_index()

pernod_posi.columns = ['EQUIPES', 'Clientes Distintos']

# Preencha valores ausentes com 0
positivacoes_completo['Positivas'] = positivacoes_completo['Positivas'].fillna(0)




# Salvar os DataFrames em um arquivo Excel com planilhas separadas
diretorio = r'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\SCRIPITADOS'

# Nome do arquivo Excel
nome_arquivo = os.path.join(diretorio, 'FASEAMENTO_PERNOD.xlsx')

# Organizando 

result_completo = result_completo.sort_values(by='Nome Equipe')
result_completo_nacional = result_completo_nacional.sort_values(by='Nome Equipe')
positivacoes_completo = positivacoes_completo.sort_values(by='Nome Equipe')
pernod_posi = pernod_posi.sort_values(by='EQUIPES')
total_equipes = total_equipes.sort_values(by='Nome Equipe')


# Salvando os DataFrames em um arquivo Excel com múltiplas planilhas
with pd.ExcelWriter(nome_arquivo) as writer:
    result_completo.to_excel(writer, sheet_name='volume_marca', index=False)
    result_completo_nacional.to_excel(writer, sheet_name='volume_nacional', index=False)
    positivacoes_completo.to_excel(writer, sheet_name='positivacoes_pernod', index=False)
    pernod_posi.to_excel(writer, sheet_name='pernod_posi', index=False)
    total_equipes.to_excel(writer, sheet_name='Volumes_completos_pernod', index=False)


