import json
import requests
from time import sleep
import subprocess
import zipfile
import os
import sys
from pathlib import Path 


#informação da versão
response = requests.get('https://api.github.com/repos/SethiCoel/atualizacao-bot-wpp/releases/latest')
latest_release = json.loads(response.text)
ultima_versao = latest_release['tag_name']

nome_do_arquivo = latest_release['assets'][0]['name']
download_url = latest_release['assets'][0]['browser_download_url']


def versao():

    with open('versao.txt', 'r', encoding='utf-8') as arquivo:
        texto = arquivo.read()
        
    return texto

def download_file():
    try:

        if ultima_versao != versao():

            print('Baixando nova versão...')
            sleep(1)

            print('Atualizando programa...')

            response = requests.get(download_url)
            with open(nome_do_arquivo, 'wb') as f:
                f.write(response.content)

            print('Atualização Finalizada!')
            
    except Exception as error:
        print(error)
        input()
    

def extrair_arquivo():
    try:
        diretorio_atual = os.getcwd()
        diretorio_anterior = os.path.dirname(diretorio_atual)

        arquivo = Path(nome_do_arquivo)

        if arquivo.exists():

            with zipfile.ZipFile(nome_do_arquivo, 'r') as zip_ref:
                zip_ref.extractall(diretorio_atual)
            
            caminho_arquivo = f'{diretorio_atual}/{nome_do_arquivo}'
            

            sleep(2)
            os.remove(caminho_arquivo)

    except Exception as error:
        print(error)
        input()


def abrir_programa():
    sleep(1)

    caminho_executavel = f'Mensagem.Automatica.exe'

    command = f'start {caminho_executavel}'
    subprocess.Popen(command, shell=True)
    sys.exit()


if __name__ == '__main__':
    download_file()
    extrair_arquivo()
    # abrir_programa()

