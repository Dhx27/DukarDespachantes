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


caminho_planilha = r'M:\TI\ROBOS\ROBOS_EM_DEV\Automação Python\DukarDespachantes\data\BASE DETRAN GOIAS.xlsx'

#Abri a planilha do excel
planilha = load_workbook(caminho_planilha)

#Passa a instacia da planilha BASE
guia_dados = planilha['BASE']

#Passa os cabeçalhos
guia_dados['A1'] = "PLACA"
guia_dados['B1'] = "RENAVAM"
guia_dados['D1'] = "STATUS"

index = 0
linhas = list(guia_dados.iter_rows(min_row=2 , max_row= guia_dados.max_row))


# Configurar o Chrome com um User-Agent falso usando undetected-chromedriver
chrome_options = Options()
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")

# Inicializa o navegador com as opções e com o stealth mode ativado
navegador = uc.Chrome(options=chrome_options)
navegador.maximize_window()

navegador.get("https://www.go.gov.br/servicos/servico/consultar-veiculo--ipva-multas-e-crlv")

#Aguarda até o elemento ficar visível (com um timeout de 60 segundos)
try:
    
    #Esperar o botão de acessas e clicar nele
    botaoAcessar = WebDriverWait(navegador, 1000).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-main/div/div/app-sidenav/div/div/div/app-servico-detalhe/div/div[2]/div/app-botao-acessar/div/button"))
    ).click()
    
    #Esperar tela de atenção para clicar no botão entrar com o GOV
    botaoAcessarGOV = WebDriverWait(navegador, 60).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/app-root/app-modal-item[1]/div/div/div/div[2]/div/div/button"))
    ).click()
    
    time.sleep(5)
    
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

    telaInicial = WebDriverWait(navegador, 60).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "exui-titulo-servico"))
    )

    botaoRealizarConsulta = WebDriverWait(navegador, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "body  exui-button-secondary > button"))
    ).click()

    while index < len(linhas):
        row = linhas[index]

        placa_atual = row[0].value
        renavam_atual = row[1].value
        status_atual = row[2].value

        if status_atual is None:

            tela_pesquisa = WebDriverWait(navegador, 60).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "exui-card-formulario-status > mat-card"))
            )

            campo_placa = navegador.find_element(By.CSS_SELECTOR, "#mat-input-0")
            campo_placa.send_keys(placa_atual)

            time.sleep(2)

            campo_renavam = navegador.find_element(By.CSS_SELECTOR, "#mat-input-1")
            campo_renavam.send_keys(renavam_atual)

            time.sleep(1.5)

            botao_consultar = navegador.find_element(By.CSS_SELECTOR, "exui-button-primary > button")
            botao_consultar.click()

            tela_dados_veiculo = WebDriverWait(navegador, 60).until(
                EC.text_to_be_present_in_element((By.CSS_SELECTOR, "body > div > div > div > div > div > lib-detalhes-veiculo > div > exui-abas > div > div > exui-aba:nth-child(1) > div > lib-dados-veiculo > div > div > exui-card-detalhamento > exui-card > mat-card > div > div"), placa_atual)
            )

            




    
        
    
    
except TimeoutException:
    print("LOGIN NO SITE NN REALIZADO")


navegador.quit()
