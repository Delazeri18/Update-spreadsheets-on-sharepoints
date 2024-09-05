import pandas as pd
import os 

#arquivo = planilha
# Defina o caminho relativo para a pasta "planilhas" dentro do projeto
caminho_projeto = os.path.dirname(os.path.abspath(__file__))
caminho_planilhas = os.path.join(caminho_projeto, 'planilhas')

# Nome do arquivo que você deseja ler
nome_arquivo = 'TEMPLATE_KA_MARIA.xlsx'

# Caminho completo para o arquivo dentro da pasta "planilhas"
caminho_arquivo = os.path.join(caminho_planilhas, nome_arquivo)

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo)
df_sem_primeira_linha = df.drop(index=0)
novos_nomes_colunas = df_sem_primeira_linha.iloc[0]
df_sem_primeira_linha.columns = novos_nomes_colunas
df_final = df_sem_primeira_linha[1:].reset_index(drop=True)

# mapeando colunas: 
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
    8815: "OUTROS",
    657: "PREMIUM",
    6195: "POPULAR CAMPARI",
    2522:"OUTROS",
    10063: "OUTROS",
    10064: "OUTROS",
    10065: "OUTROS",
    3423: 'OUTROS',
    539: "PREMIUM",
    925: 'PREMIUM',
    2503 : 'PREMIUM',
    2804 : 'OUTROS',
    832 : 'PREMIUM',
    2520: 'OUTROS',
    1522 : 'PREMIUM',
    3426 : 'OUTROS',
    8733 : 'PREMIUM',
    1399 : 'PREMIUM',
    1398 : 'PREMIUM',
    4327 : 'PREMIUM',
    2523 : "PREMIUM",
    2521 : "PREMIUM",
    716 : 'OUTROS',
    715 : 'OUTROS',
    8825 : 'PREMIUM',
    693: 'POPULAR CAMPARI',
    2522 : 'POPULAR CAMPARI',
    388 : 'POPULAR CAMPARI',
    6668 : "PREMIUM",
    4588 : "PREMIUM",
    5265 : "PREMIUM",
    4456 : "PREMIUM",
    4457 : "PREMIUM",
    3653 : 'PREMIUM'
}

df_final['MARCA'] = df_final['Produto'].map(mapeamento_produtos)


quantidade = df_final.groupby('MARCA')['Qtde'].sum().reset_index()

# Salvar os DataFrames em um arquivo Excel com planilhas separadas
diretorio = r'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\SCRIPITADOS'

# Nome do arquivo Excel
nome_arquivo = os.path.join(diretorio, 'CAMPARI_MARIA_RITA.xlsx')

# Salvando os DataFrames em um arquivo Excel com múltiplas planilhas
with pd.ExcelWriter(nome_arquivo) as writer:
    quantidade.to_excel(writer, sheet_name='df_posi_PR', index=False)