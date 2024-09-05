import pandas as pd
import os 

# Defina o caminho relativo para a pasta "planilhas" dentro do projeto
caminho_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_planilhas = os.path.join(caminho_projeto, 'planilhas')

# Nome do arquivo que você deseja ler
nome_arquivo = 'template_pernod.xlsx'

# Caminho completo para o arquivo dentro da pasta "planilhas"
caminho_arquivo = os.path.join(caminho_planilhas, nome_arquivo)

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo)
df_sem_primeira_linha = df.drop(index=0)
novos_nomes_colunas = df_sem_primeira_linha.iloc[0]
df_sem_primeira_linha.columns = novos_nomes_colunas
df_final = df_sem_primeira_linha[1:].reset_index(drop=True)
# REALAIZANDO MAPEAMENTO

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
    "EVENTOS PR": "Nao",
    "KEY ACCOUNT SC ON": "ON_SC"
}

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

# Arrumando DF
df_final['EQUIPES'] = df_final['Nome Equipe'].map(mapeamento_equipes)
df_final['MARCA'] = df_final['Produto'].map(mapeamento)
df_final['LOCAL'] = df_final['MARCA'].apply(classificar_categoria)

# convertendo filial para int
df_final['Filial']=df_final['Filial'].astype(int)
df_final['Cliente'] = df_final['Cliente'].astype(int)

# retirando linhas de enventos
df_filtrado = df_final.drop(df_final[df_final['Nome Equipe'] == 'EVENTOS PR'].index)

# Resetando o índice do DataFrame resultante
df_filtrado.reset_index(drop=True, inplace=True)

# separando filiais
df_filial_PR = df_filtrado[df_filtrado['Filial'] != 6]
df_filial_SC = df_filtrado[df_filtrado['Filial'] == 6 ]


# Volume

volumes_PR = df_filial_PR.groupby('MARCA')['Volumes'].sum().reset_index()
volumes_SC = df_filial_SC.groupby('MARCA')['Volumes'].sum().reset_index()


marcas_interesse = ['ABSOLUT', 'BALLANTINES', 'BEEFEATER', 'CHIVAS', 'SAO FRANCISCO', 'DOMECQ', 'PASSPORT', 'ORLOFF', 'NATU']
todas_marcas = pd.DataFrame({'MARCA': marcas_interesse})
volumes_pr = pd.merge(todas_marcas, volumes_PR, on='MARCA', how='left')
volumes_sc = pd.merge(todas_marcas, volumes_SC, on='MARCA', how='left')

volumes_nacional_SC = df_filial_SC.groupby('LOCAL')['Volumes'].sum().reset_index()
volumes_nacional_PR = df_filial_PR.groupby('LOCAL')['Volumes'].sum().reset_index()


volumes_pr['Volumes'] = volumes_pr['Volumes'].fillna(0)
volumes_sc['Volumes'] = volumes_sc['Volumes'].fillna(0)


volumes_pr_sorted = volumes_pr.sort_values(by='MARCA').reset_index(drop=True)
volumes_sc_sorted = volumes_sc.sort_values(by='MARCA').reset_index(drop=True)

# contando positivações
df_filial_PR_unicos = df_filial_PR.drop_duplicates(subset='Cliente')
df_filial_SC_unicos = df_filial_SC.drop_duplicates(subset='Cliente')

# Contar número de clientes únicos (positivações) em cada filial
positivacao_PR = df_filial_PR_unicos['Cliente'].count()
positivacao_SC = df_filial_SC_unicos['Cliente'].count()

dados_positivacao = {
    'ESTADO': ['PR','SC'],
    'POSITIVACAO' : [positivacao_PR, positivacao_SC]
}

df_positivacao = pd.DataFrame(dados_positivacao)

# Salvar os DataFrames em um arquivo Excel com planilhas separadas
diretorio = r'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\SCRIPITADOS'

# Nome do arquivo Excel
nome_arquivo = os.path.join(diretorio, 'PERNOD_ACOMPANHAMENTO.xlsx')

# Salvando os DataFrames em um arquivo Excel com múltiplas planilhas
with pd.ExcelWriter(nome_arquivo) as writer:
    volumes_pr_sorted.to_excel(writer, sheet_name='df_volume_PR', index=False)
    df_positivacao.to_excel(writer, sheet_name='positivacao', index=False)
    volumes_sc_sorted.to_excel(writer, sheet_name='volume_SC', index=False)
    volumes_nacional_PR.to_excel(writer, sheet_name='volumes_nacional_PR', index=False)
    volumes_nacional_SC.to_excel(writer, sheet_name='volumes_nacional_SC', index=False)
