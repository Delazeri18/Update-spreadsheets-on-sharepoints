import pandas as pd
import os 

# Defina o caminho relativo para a pasta "planilhas" dentro do projeto
caminho_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_planilhas = os.path.join(caminho_projeto, 'planilhas')

# Nome do arquivo que você deseja ler
nome_arquivo = 'VCT_PR_ONN_OF.xlsx'

# Caminho completo para o arquivo dentro da pasta "planilhas"
caminho_arquivo = os.path.join(caminho_planilhas, nome_arquivo)

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo)
df_sem_primeira_linha = df.drop(index=0)
novos_nomes_colunas = df_sem_primeira_linha.iloc[0]
df_sem_primeira_linha.columns = novos_nomes_colunas
df_final = df_sem_primeira_linha[1:].reset_index(drop=True)

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
    "HEINEKEN": "MC",
    "TELEVENDAS": "MC",
    "EVENTOS PR": "MC"
}

mapeamento_classes = {
    1381: "CELLAR",
    1573: "DEVILS",
    1575: "MARQUES",
    1901: "RESERVADO",
    1902: "RESERVADO",
    1903: "RESERVADO",
    1904: "RESERVADO",
    1905: "RESERVADO",
    1909: "GRAN RESERVA",
    1910: "CASILLERO",
    1911: "CASILLERO",
    1912: "CASILLERO",
    1913: "CASILLERO",
    1914: "CASILLERO",
    1915: "CASILLERO",
    1916: "RESERVADO",
    1918: "RESERVADO",
    1923: "MARQUES",
    1934: "MARQUES",
    1935: "MARQUES",
    1936: "MARQUES",
    1937: "DIABLO",
    2001: "TRIVENTO",
    2002: "TRIVENTO",
    2539: "CELLAR",
    2947: "CELLAR",
    2994: "RESERVADO",
    3039: "CASILLERO",
    3199: "MARQUES",
    3252: "MARQUES",
    3275: "DIABLO",
    3286: "CELLAR",
    3288: "CELLAR",
    3291: "MARQUES",
    3292: "CELLAR",
    3293: "CELLAR",
    3303: "GRAN RESERVA",
    3304: "CELLAR",
    3424: "CASILLERO",
    3425: "CASILLERO",
    3442: "TRIVENTO",
    3458: "CELLAR",
    3511: "CASILLERO",
    3512: "CASILLERO",
    3513: "CASILLERO",
    3602: "MARQUES",
    3620: "CELLAR",
    3702: "CELLAR",
    3703: "CELLAR",
    3704: "TRIVENTO",
    3754: "CASILLERO",
    3996: "CELLAR",
    4149: "DIABLO",
    4344: "CASILLERO",
    4345: "CASILLERO",
    4802: "CELLAR",
    4999: "TRIVENTO",
    5007: "RESERVADO",
    5043: "CELLAR",
    5073: "CELLAR",
    5074: "CELLAR",
    5075: "CELLAR",
    5076: "CELLAR",
    5077: "CELLAR",
    5078: "CELLAR",
    5079: "CELLAR",
    5080: "CELLAR",
    5081: "MARQUES",
    6787: "MARQUES",
    8719: "TRIVENTO",
    8720: "GRAN RESERVA",
    8926: "CELLAR",
    9525: "CELLAR",
    9526: "CELLAR",
    9527: "TRIVENTO",
    9528: "CELLAR",
    9529: "RESERVADO",
    9598: "DIABLO",
    9599: "CDD CARNIVAL",
    9600: "CDD CARNIVAL",
    9601: "CDD CARNIVAL",
    9602: "CDD CARNIVAL",
    9603: "CELLAR",
    9617: "CDD BELIGHT",
    9671: "CDD BELIGHT",
    9701: "RESERVADO",
    9825: "MARQUES",
    9826: "CELLAR",
    9827: "CELLAR",
    3278: "CASILLERO"
}

df_final['CLASSE'] = df_final['Produto'].map(mapeamento_classes)
df_final['EQUIPE'] = df_final['Nome Equipe'].map(mapeamento_equipes)

# Lista das marcas de interesse
marcas_interesse = ["CASILLERO", "TRIVENTO", "RESERVADO", "DIABLO", "MARQUES"]

# Agrupar por Marca e somar os volumes
volumes_por_marca = df_final.groupby('CLASSE')['Volumes'].sum().reset_index()

#Criar DataFrame com todas as marcas de interesse
todas_marcas = pd.DataFrame({'CLASSE': marcas_interesse})

# Combinar os resultados com todas as marcas de interesse
volumes_completo = pd.merge(todas_marcas, volumes_por_marca, on='CLASSE', how='left')

# Preencher valores ausentes com 0
volumes_completo['Volumes'] = volumes_completo['Volumes'].fillna(0)

volumes_completo_sorted = volumes_completo.sort_values(by='CLASSE').reset_index(drop=True)

# POSITIVACOES
positivacao = df_final.groupby('CLASSE')['Cliente'].nunique().reset_index()
marcas_interesse2 = ["CASILLERO", "TRIVENTO", "RESERVADO", "DIABLO", "MARQUES",'CDD CARNIVAL','CDD BELIGHT']


todas_marcas2 = pd.DataFrame({'CLASSE': marcas_interesse2})

positivacao_completo = pd.merge(todas_marcas2, positivacao, on='CLASSE', how='left')

# Preencher valores ausentes com 0
positivacao_completo['Cliente'] = positivacao_completo['Cliente'].fillna(0)

positivacao_completo = positivacao_completo.sort_values(by='CLASSE').reset_index(drop=True)


# Salvar os DataFrames em um arquivo Excel com planilhas separadas
diretorio = r'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\SCRIPITADOS'

# Nome do arquivo Excel
nome_arquivo = os.path.join(diretorio, 'Acompanhamento_VCT.xlsx')


with pd.ExcelWriter(nome_arquivo) as writer:
    positivacao_completo.to_excel(writer, sheet_name='POSITIVA', index=False)
    volumes_completo_sorted.to_excel(writer, sheet_name='VOLUME', index=False)