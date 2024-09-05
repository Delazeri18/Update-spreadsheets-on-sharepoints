import os
import subprocess

# Caminho da pasta que contém os arquivos Excel
caminho_pasta_planilhas = 'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS\\planilhas'

# Caminho da pasta que contém os scripts Python
caminho_pasta_scripts = 'C:\\Users\\Kewin Delazeri\\Documents\\SCRIPT_ACOMPANHAMENTOS'

# Dicionário para mapear arquivos Excel para scripts Python
scripts = {
    'Campari_equipe.xlsx': 'FASEAMENTO_CAMPARI.py',
    'Campari_p4p.xlsx': 'CAMPARI_p4p.py',
    'FASEAMENTO_VCT.xlsx': 'FASEMAENTO_VCT.py',
    'PERNOD_EQUIPES.xlsx': 'FASEAMENTO_PERNOD.py',
    'TEMPLATE_BEAM_SUNTORY.xlsx': 'BEAM.py',
    'TEMPLATE_JACK.xlsx': 'JACK.py',
    'template_pernod.xlsx': 'PERNOD.py',
    'VCT_PR_ONN_OF.xlsx': 'VCT_ONN_OF.py',
    'TEMPLATE_KA_MARIA.xlsx': 'MARIA_RITA.py',
    'Template_CC.xlsx' : 'CAMPARI_CLUB.py',
    'Equipes_jack.xlsx' : 'faseamento_JACK',
    'Equipe_Beam.xlsx' : 'faseamento_beam.py'
}

def executar_script(nome_arquivo):
    if nome_arquivo in scripts:
        script = scripts[nome_arquivo]
        caminho_script = os.path.join(caminho_pasta_scripts, script)
        print(f'Executando {caminho_script}...')
        try:
            subprocess.run(['python', caminho_script], check=True)
        except subprocess.CalledProcessError as e:
            print(f'Erro ao executar {caminho_script}: {e}')

# Iterar sobre os arquivos na pasta de planilhas
for nome_arquivo in os.listdir(caminho_pasta_planilhas):
    caminho_completo = os.path.join(caminho_pasta_planilhas, nome_arquivo)
    
    if os.path.isfile(caminho_completo):
        executar_script(nome_arquivo)

# Executar o script final
caminho_script_final = os.path.join(caminho_pasta_scripts, 'Mandar_share.py')
print('Executando Mandar_share.py...')
try:
    subprocess.run(['python', caminho_script_final], check=True)
except subprocess.CalledProcessError as e:
    print(f'Erro ao executar {caminho_script_final}: {e}')
