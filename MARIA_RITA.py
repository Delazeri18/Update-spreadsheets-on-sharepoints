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
    4339: "PREMIUM",
    6195: "POPULAR CAMPARI",
    10063: "PREMIUM",
    10064: "PREMIUM",
    10065: "PREMIUM",
    925: 'PREMIUM',
    2503 : 'PREMIUM',
    2804 : 'PREMIUM',
    832 : 'PREMIUM',
    2520: 'PREMIUM',
    1522 : 'PREMIUM',
    8733 : 'PREMIUM',
    1399 : 'PREMIUM',
    1398 : 'PREMIUM',
    4327 : 'CAMPARI',
    2523 : "PREMIUM",
    2521 : "PREMIUM",
    693: 'POPULAR CAMPARI',
    2522 : 'PREMIUM',
    388 : 'POPULAR CAMPARI',
    4456 : "PREMIUM",
    3653 : 'PREMIUM'
}

def classificar_categoria(marca):
    return mapeamento_produtos.get(marca, "OUTROS")


df_final['MARCA'] = df_final['Produto'].apply(classificar_categoria)


quantidade = df_final.groupby('MARCA')['Qtde'].sum().reset_index()

# Salvar os DataFrames em um arquivo Excel com planilhas separadas
diretorio = r'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\SCRIPITADOS'

# Nome do arquivo Excel
nome_arquivo = os.path.join(diretorio, 'CAMPARI_MARIA_RITA.xlsx')

# Salvando os DataFrames em um arquivo Excel com múltiplas planilhas
with pd.ExcelWriter(nome_arquivo) as writer:
    quantidade.to_excel(writer, sheet_name='df_posi_PR', index=False)