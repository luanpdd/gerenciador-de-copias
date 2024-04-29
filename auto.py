import os
import shutil

def copiar_arquivos_xlsx(pasta_origem):
    # Criar a pasta destino dentro da pasta de origem
    pasta_destino = os.path.join(pasta_origem, "arquivos do excel")
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
    
    # Percorrer as subpastas e arquivos
    for diretorio, subpastas, arquivos in os.walk(pasta_origem):
        for arquivo in arquivos:
            if arquivo.endswith(".xlsx"):
                caminho_completo = os.path.join(diretorio, arquivo)
                shutil.copy(caminho_completo, pasta_destino)
                print(f"Arquivo {arquivo} copiado para {pasta_destino}")

# Entrada do usuário para o caminho da pasta de origem
pasta_origem = input("Digite o caminho da pasta de origem: ")
copiar_arquivos_xlsx(pasta_origem)
