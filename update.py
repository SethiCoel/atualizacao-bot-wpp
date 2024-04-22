import json
import requests
from time import sleep
import subprocess
import zipfile
import os
import sys

def atualizar_programa():
    try:
        response = requests.get('https://api.github.com/repos/SethiCoel/atualizacao-bot-wpp/releases/latest')
        latest_release = json.loads(response.text)
        ultima_versao = latest_release['tag_name']

        print('Nova versão encontrada!')
        sleep(1)
        nome_do_arquivo = latest_release['assets'][0]['name']
        download_url = latest_release['assets'][0]['browser_download_url']
        download_file(download_url, nome_do_arquivo)

    except Exception as error:
        print(error)

def download_file(url, nome):
    print('Baixando nova versão...')
    sleep(1)

    print('Atualizando programa...')

    diretorio_atual = os.getcwd()
    response = requests.get(url)
    with open(nome, 'wb') as f:
        f.write(response.content)

    print('Atualização Finalizada!')

    extrair_arquivo(nome, diretorio_atual)

def extrair_arquivo(arquivo_zip, extrair_para):
    with zipfile.ZipFile(arquivo_zip, 'r') as zip_ref:
        zip_ref.extractall(extrair_para)

def abrir_programa():
    caminho_executavel = f'Mensagem.Automatica.exe'
    command = f'start {caminho_executavel}'
    subprocess.Popen(command, shell=True)
    sys.exit()

if __name__ == '__main__':
    atualizar_programa()
    abrir_programa()

