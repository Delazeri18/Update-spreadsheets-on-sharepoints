import os
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

# Configurações
site_url = "https://grupolorac.sharepoint.com/sites/Dinho"
username = "kewin.delazer@grupolorac.onmicrosoft.com"
password = '@@Bebidas2024'
sharepoint_folder = "/sites/Dinho/Documentos Compartilhados/Dados/Comercial/Equipes/Gerencial/3 TRI 2024/Análises Comercial"
local_folder = "C:/Users/Kewin Delazeri/Documents/SCRIPT_ACOMPANHAMENTOS/SCRIPITADOS"

# Autenticação
ctx_auth = AuthenticationContext(site_url)
if not ctx_auth.acquire_token_for_user(username, password):
    print("Autenticação falhou")
    exit()

ctx = ClientContext(site_url, ctx_auth)

# Envio dos Arquivos
for file_name in os.listdir(local_folder):
    file_path = os.path.join(local_folder, file_name)
    
    if os.path.isfile(file_path):
        with open(file_path, "rb") as file_content:
            target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
            target_file = target_folder.upload_file(file_name, file_content)
            ctx.execute_query()
            print(f"Arquivo '{file_name}' enviado com sucesso!")
    else:
        print(f"'{file_name}' não é um arquivo. Ignorado.")

print("Todos os arquivos foram enviados com sucesso!")
