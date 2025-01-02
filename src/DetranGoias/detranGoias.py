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
    campoCPF.send_keys(os.getenv("CPF"))

    botaoContinuarLogin = navegador.find_element(By.CSS_SELECTOR, "#enter-account-id")
    botaoContinuarLogin.click()

    campoSenha =  WebDriverWait(navegador, 1000).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "#password"))
    )
    campoSenha.send_keys(os.getenv("SENHA"))

    botaoEntrar = navegador.find_element(By.CSS_SELECTOR, "#submit-button")
    botaoEntrar.click()
    
    print ("DIogo")
    
    
except TimeoutException:
    print("LOGIN NO SITE NN REALIZADO")


navegador.quit()
