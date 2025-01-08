#ROBO DE EMISSÃO DE IPVA E LICENCIAMENTO

from dotenv import load_dotenv
from openpyxl import load_workbook
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import undetected_chromedriver as uc
import requests
from selenium.webdriver.common.keys import Keys  # Para acessar as teclas especiais
from selenium.webdriver.common.action_chains import ActionChains
import os
import pyautogui
import pdfplumber
import re

load_dotenv()

pasta_download = r''
caminho_planilha = r''

# Dados do Turnstile CAPTCHA
API_KEY = os.getenv("api_key")
SITEKEY = "0x4AAAAAAAO9omZCUc8pnQfN"
PAGE_URL = "https://internet.detrannet.mt.gov.br/ConsultaVeiculo.asp"

# 1. Enviar requisição para resolver o CAPTCHA
def enviar_requisicao_captcha(api_key, sitekey, page_url):
    url = "http://2captcha.com/in.php"
    data = {
        "key": api_key,
        "method": "turnstile",
        "sitekey": sitekey,
        "pageurl": page_url,
    }
    response = requests.post(url, data=data)
    if response.status_code == 200 and "OK|" in response.text:
        captcha_id = response.text.split('|')[1]
        return captcha_id
    else:
        raise Exception(f"Erro ao enviar captcha: {response.text}")

# 2. Obter o token de resposta do CAPTCHA
def obter_resposta_captcha(api_key, captcha_id):
    url = f"http://2captcha.com/res.php?key={api_key}&action=get&id={captcha_id}"
    while True:
        response = requests.get(url)
        if "CAPCHA_NOT_READY" in response.text:
            print("Captcha ainda não resolvido, tentando novamente em 5 segundos...")
            time.sleep(5)
        elif "OK|" in response.text:
            return response.text.split('|')[1]
        else:
            raise Exception(f"Erro ao obter resposta do captcha: {response.text}")

# 3. Inserir o token na página (use Selenium ou outro método)
def inserir_token(navegador, token):
    try:

        navegador.execute_script(
            "document.getElementById('g-recaptcha-response').style.display = 'block';"
        )
        navegador.execute_script(
            f"document.getElementById('g-recaptcha-response').innerHTML = '{token}';"
        )
        navegador.execute_script(
            "document.getElementsByName('g-recaptcha-response')[0].dispatchEvent(new Event('change'));")
        print("Token inserido com sucesso!")

    except Exception as e:
        raise Exception(f"Erro ao inserir token: {str(e)}")



#Cria uma pasta de saída baseada no nome da planilha e em um caminho de base fornecido.
def criar_pasta_saida(caminho_planilha, pasta_downloads):

    #Pega o nome do arquivo excel sem a extensão
    nome_arquivo = os.path.splitext(os.path.basename(caminho_planilha))[0]

    #Caminho da pasta saida com base na pasta fornecida
    pasta_saida = os.path.join(pasta_download, nome_arquivo)

    #Verifica se a pasta de saída já existe, se não existir, cria a pasta
    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)
        print(f"Pasta criada: {pasta_saida}")
    else:
        print(f"A pasta já está criada {pasta_saida}")


# Configurar o Chrome com um User-Agent falso usando undetected-chromedriver
chrome_options = Options()
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")

# Inicializa o navegador com as opções e com o stealth mode ativado
navegador = uc.Chrome(options=chrome_options)
navegador.maximize_window()

navegador.get("https://www.detran.mt.gov.br/")




navegador.quit()
