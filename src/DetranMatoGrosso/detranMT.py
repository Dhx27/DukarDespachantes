#ROBO DE EMISSÃO DE LICENCIAMENTO E IPVA

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
import pyautogui

load_dotenv()

pasta_download = r'C:\Users\Diogo Lana\Desktop\Nova pasta'
caminho_planilha = r'C:\Users\Diogo Lana\Desktop\Nova pasta\MODELO MT.xlsx'

# Dados do Turnstile CAPTCHA
API_KEY = os.getenv("api_key")
SITEKEY = "0x4AAAAAAAO9omZCUc8pnQfN"
PAGE_URL = "https://internet.detrannet.mt.gov.br/ConsultaVeiculo.asp"


def selecionar_download_como_pdf_lic(pasta_saida, placa_atual):
    
    time.sleep(5)
    
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(3)
    
    # Navegar pelas opções até "Salvar"
    for _ in range(5):  # Pressiona 'tab' 5 vezes
        pyautogui.hotkey('tab')
    time.sleep(1.5)

    pyautogui.write("salvar")  # Digita "salvar"
    time.sleep(1.5)

        # Navegar até o botão de salvar
    for _ in range(4):  # Pressiona 'tab' 6 vezes
        pyautogui.hotkey('tab')
    time.sleep(1.5)

    pyautogui.hotkey('enter')  # Confirma "Salvar"
    time.sleep(3)

    # Define o caminho completo para salvar o arquivo
    caminho_download = os.path.join(pasta_saida, f"LIC {placa_atual}")
    caminho_download = os.path.normpath(caminho_download)  # Normaliza o caminho

    # Digitar o caminho de salvamento
    pyautogui.write(caminho_download)
    time.sleep(1)
    pyautogui.hotkey('enter')  # Confirma o caminho

    print("Arquivo salvo com sucesso!")
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(1.5)
    pyautogui.hotkey('ctrl', 'w')
    
    

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

        elemento = WebDriverWait(navegador, 30).until(
            EC.presence_of_element_located((By.XPATH, "//input[@name='g-recaptcha-response']"))
        )
        
        valor_captcha = f"{token}"
        navegador.execute_script(f"arguments[0].setAttribute('value', '{valor_captcha}')", elemento)

        # Disparar o evento de mudança (opcional)
        #navegador.execute_script("arguments[0].dispatchEvent(new Event('change'))", elemento)



    except Exception as e:
        raise Exception(f"Erro ao inserir token: {str(e)}")



#Cria uma pasta de saída baseada no nome da planilha e em um caminho de base fornecido.
def criar_pasta_saida(pasta_download, caminho_planilha):

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

    return pasta_saida

# Configurar o Chrome com um User-Agent falso usando undetected-chromedriver
chrome_options = Options()
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")

# Inicializa o navegador com as opções e com o stealth mode ativado
navegador = uc.Chrome(options=chrome_options)
navegador.maximize_window()

pasta_saida = criar_pasta_saida(pasta_download, caminho_planilha)

#Abri a planilha do excel
planilha = load_workbook(caminho_planilha)

#Passa a instacia da planilha BASE
guia_dados = planilha['BASE']

#                           EMISSÃO LICENCIAMENTO 

navegador.get("https://www.detran.mt.gov.br/")

tela_atendimento = WebDriverWait(navegador, 100).until(

    EC.visibility_of_element_located((By.CSS_SELECTOR, "#myPopup > a > picture > img"))
)
        
botao_fechar = navegador.find_element(By.CSS_SELECTOR, "#myPopup > span")
navegador.execute_script("arguments[0].click();", botao_fechar)
        
tela_cookies = WebDriverWait(navegador, 30).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="adopt-accept-all-button"]'))
).click()
        

# Pegar todas as guias abertas
abas = navegador.window_handles


index = 0
linhas = list(guia_dados.iter_rows(min_row=2, max_row=guia_dados.max_row))
linhaPlan = 1

'''
while index < len(linhas):
    linhaPlan += 1

    row = linhas[index]

    #Obtem todos os valores dispobniveis na planilha
    placa_atual = row[0].value
    renavam_atual = row[1].value
    cnpj_atual = row[2].value
    status_lic = row[3].value

    if status_lic is None:
        
        campo_placa_renavam = WebDriverWait(navegador, 100).until(
            EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div[4]/section/div/div/div/div/div[1]/section/div/div/div/div/div/div[2]/div/div/section/div/div[2]/div/div/div/div/div'))
        )
        
        campo_placa = navegador.find_element(By.CSS_SELECTOR, "#input_placa")
        campo_placa.clear()
        campo_placa.send_keys(placa_atual)
        
        time.sleep(1.5)
        
        campo_renavam = navegador.find_element(By.CSS_SELECTOR, "#input_renavam")
        campo_renavam.clear()
        campo_renavam.send_keys(renavam_atual)
        
        botao_consultar = WebDriverWait(navegador, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#formVeiculo > div:nth-child(4) > input.dtrn-frm-sub.dtrn-vin.text-size-acessibilidade"))
        ).click()
        
        # Pegar todas as guias abertas
        abas = navegador.window_handles
        
        # Alternar para a nova guia (última aberta)
        navegador.switch_to.window(abas[-1])
        print("Estamos na nova guia:", navegador.title)
        
        tela_cnpj = WebDriverWait(navegador, 60).until(
            EC.visibility_of_element_located((By.XPATH, '/html/body/center/div/form/table/tbody/tr[3]/td[2]/input'))
        )

        time.sleep(5)
        
        try:
            print("Enviando requisição para resolver o CAPTCHA...")
            captcha_id = enviar_requisicao_captcha(API_KEY, SITEKEY, PAGE_URL)
            print("Obtendo resposta do CAPTCHA...")
            token = obter_resposta_captcha(API_KEY, captcha_id)
            print("Solução do CAPTCHA obtida:")
            inserir_token(navegador, token)  # Descomente ao integrar com Selenium
            
            cnpj = cnpj_atual.replace(".", "").replace("/", "").replace("-", "")
            
            campo_cnpj = navegador.find_element(By.XPATH, '/html/body/center/div/form/table/tbody/tr[3]/td[2]/input')
            campo_cnpj.send_keys(cnpj)
            
            botao_consultar_veiculo = WebDriverWait(navegador, 60).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/center/div/form/table/tbody/tr[4]/td/input'))
            ).click()
            
            tela_baixar_boleto = WebDriverWait(navegador, 60).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "#td_meio_Integral"))
            )
            
            modal_licenciamento = navegador.find_element(By.CSS_SELECTOR, "#cmbTipoDebito")
            modal_licenciamento.click()
            
            botao_download_guia = WebDriverWait(navegador, 60).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '#spanDAR_LicenciamentoExercicio))
            ).click()
            
            selecionar_download_como_pdf_lic(pasta_saida, placa_atual)
            
            guia_dados[f'D{linhaPlan}'] = "OK!"  
            planilha.save(caminho_planilha)
            
            # Voltar para a guia original
            navegador.switch_to.window(abas[0])
            print("Voltamos para a guia original:", navegador.title)
            
            index += 1
            
        except Exception as e:
            print(f"Erro: {str(e)}")
    else:
        index += 1
        continue
'''

#                                       EMISSÃO IPVA   
 
navegador.get("https://www.sefaz.mt.gov.br/ipva/emissaoguia/emitir")


index = 0
linhas = list(guia_dados.iter_rows(min_row=2, max_row=guia_dados.max_row))
linhaPlan = 1

while index < len(linhas):
    linhaPlan += 1

    row = linhas[index]
    
    renavam_atual_ipva = row[1].value
    status_atual_ipva = row[4].value
    
    if status_atual_ipva is None:
        
        tela_pesquisa_ipva = WebDriverWait(navegador, 60).until(
            EC.visibility_of_element_located((By.XPATH, '/html/body/center/form/table/tbody/tr[5]/td[2]/input'))
        )
        
        campo_input_renavam = navegador.find_element(By.XPATH, '/html/body/center/form/table/tbody/tr[5]/td[2]/input')
        campo_input_renavam.clear()
        campo_input_renavam.send_keys(renavam_atual_ipva)
        
        botao_consultar = WebDriverWait(navegador, 60).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/center/form/p/input[1]'))
        ).click()
        
        #Aqui pedi captcha
        
        navegador.refresh()
        
        tela_ipva = WebDriverWait(navegador, 60).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "table.SEFAZ-TABLE-Moldura"))
        )
        
        elemento_desconto5 = WebDriverWait(navegador, 60).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#tipoGuia\.2025\.2"))
        ).click()
        
        botao_gerar_guias = WebDriverWait(navegador, 60).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#btnOK"))
        ).click()
        
        navegador.refresh()
        
        
    else:
        
        index += 1 
        continue
    


navegador.quit()
