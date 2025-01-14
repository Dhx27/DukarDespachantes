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

pasta_download = r'C:\Users\Robo01\Desktop\ENTRADA\SAIDA'
caminho_planilha = r'C:\Users\Robo01\Desktop\ENTRADA\MODELO DF.xlsx'

def criar_pasta_saida(pasta_download, caminho_planilha):
    
    # Pega o nome do arquivo da planilha sem a extensão
    nome_arquivo = os.path.splitext(os.path.basename(caminho_planilha))[0]

    #Caminho da pasta de saída com base na pasta fornecida
    pasta_saida = os.path.join(pasta_download, nome_arquivo)

    if not os.path.exists(pasta_saida):

        os.makedirs(pasta_saida)
        print(f"Pasta criada: {pasta_saida}")
    else:
        print(f"a pasta já existe: {pasta_saida}")

    return pasta_saida

# Caminha para selecionar o download como PDF
def selecionar_download_como_pdf(pasta_saida, placa_atual):
    
    # Navegar pelas opções até "Salvar"
    for _ in range(5):  # Pressiona 'tab' 5 vezes
        pyautogui.hotkey('tab')
    time.sleep(1.5)

    pyautogui.write("salvar")  # Digita "salvar"
    time.sleep(1.5)

        # Navegar até o botão de salvar
    for _ in range(3):  # Pressiona 'tab' 6 vezes
        pyautogui.hotkey('tab')
    time.sleep(1.5)

    pyautogui.hotkey('enter')  # Confirma "Salvar"
    time.sleep(3)

    # Define o caminho completo para salvar o arquivo
    caminho_download = os.path.join(pasta_saida, f"{placa_atual}")
    caminho_download = os.path.normpath(caminho_download)  # Normaliza o caminho

    # Digitar o caminho de salvamento
    pyautogui.write(caminho_download)
    time.sleep(1)
    pyautogui.hotkey('enter')  # Confirma o caminho

    print("Arquivo salvo com sucesso!")

pasta_saida = criar_pasta_saida(pasta_download, caminho_planilha)

#Abri a planilha do excel 
planilha = load_workbook(caminho_planilha)

#Passa a instacia da planilha BASE
guia_dados = planilha['BASE']


index = 0
linhas = list(guia_dados.iter_rows(min_row=2, max_row=guia_dados.max_row))

# Configurar o Chrome com um User-Agent falso usando undetected-chromedriver
chrome_options = Options()
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")

# Inicializa o navegador com as opções e com o stealth mode ativado
navegador = uc.Chrome(options=chrome_options)
navegador.maximize_window()

navegador.get("https://portal.detran.df.gov.br/#/home")

try:

    botao_login = WebDriverWait(navegador, 60).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-header/mat-toolbar/span/span/span[3]/button"))
    ).click()

    botao_login_gov = WebDriverWait(navegador, 60).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "#enviar"))
    ).click()
    
    campoCPF = WebDriverWait(navegador, 60).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#accountId"))
    )
    campoCPF.send_keys(os.getenv("CPF_GOIAS"))

    botaoContinuarLogin = navegador.find_element(By.CSS_SELECTOR, "#enter-account-id")
    botaoContinuarLogin.click()

    campoSenha =  WebDriverWait(navegador, 1000).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#password"))
    )
    campoSenha.send_keys(os.getenv("SENHA_GOIAS"))

    botaoEntrar = navegador.find_element(By.CSS_SELECTOR, "#submit-button")
    botaoEntrar.click()

    try:

        botao_autorizacao = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#authorize-info > div.button-panel > button.button-ok"))
        ).click()
    except Exception:
        pass

    time.sleep(5)

    tela_detran_df = WebDriverWait(navegador, 60).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "app-home > section > div:nth-child(1)"))
    )

    cabecalho_veiculos = WebDriverWait(navegador, 60).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "body > app-root > app-header > mat-toolbar > span > span > section > button:nth-child(2)"))
    ).click()

    time.sleep(3)

    botao_consulta_debitos = WebDriverWait(navegador, 60).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#headerMenu_veiculosConsultas > app-menu-item:nth-child(1) > div.mat-tooltip-trigger.menu-item-wrapper > a > article > h3"))
    ).click()

    botao_consulta_veiculos = WebDriverWait(navegador, 60).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "app-botao-consulta-veiculo-outro-proprietario > div > button"))
    ).click()

    linhaPlan =1

    while index < len(linhas):

        linhaPlan +=1
        row = linhas[index]

        placa_atual = row[0].value
        renavam_atual = row[1].value
        status_atual = row[2].value

        if status_atual is None:

            campo_placa = navegador.find_element(By.XPATH, '(//input[contains(@class, "mat-input-element")])[1]')
            campo_placa.send_keys(placa_atual)

            time.sleep(2)

            campo_renavam = navegador.find_element(By.XPATH, '(//input[contains(@class, "mat-input-element")])[2]')
            campo_renavam.send_keys(renavam_atual)

            time.sleep(1.5)

            botao_consultar = navegador.find_element(By.CSS_SELECTOR, " app-consulta-placa-renavam > mat-card-actions > div > button")
            botao_consultar.click()

            tela_dados_veiculo = WebDriverWait(navegador, 60).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "mat-card-content > app-lista-dados"))
            )

            botao_consultar_2 = navegador.find_element(By.XPATH, "/html/body/app-root/div/div[1]/app-layout-servicos/app-debitos/app-card-detran/mat-card/mat-card-content/app-lista-dados/div/div[2]/table/tbody/tr/td[7]/button")
            botao_consultar_2.click()

            tela_licenciamento = WebDriverWait(navegador, 60).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "app-lista-dados:nth-child(2) div.forms-wrapper-header.ng-star-inserted > div > span"))
            )

            

except Exception:

    print("Renicie a automação")
    breakpoint()

navegador.quit()