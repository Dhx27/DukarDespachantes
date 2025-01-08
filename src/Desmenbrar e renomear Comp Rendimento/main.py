import pdfplumber
from PyPDF2 import PdfReader, PdfWriter
import re
import os
from openpyxl import Workbook
from datetime import datetime

caminho_entrada = r'C:\Users\diogo.lana\Desktop\TESTE'

lista_pdf = [f for f in os.listdir(caminho_entrada) if f.endswith('.pdf')]

pasta_desmebrados = r'M:\FINANCEIRO\ROBO\COMPROVANTES RENDIMENTO'

#instancia a lista dos resultados dos boletos
resultados_lista = []

for pdf in lista_pdf:

    caminho_pdf = os.path.join(caminho_entrada, pdf)

    # Caminho para o arquivo de erros
    caminho_erros = os.path.join(os.path.dirname(caminho_pdf), "erros.txt")

    # Abrir o PDF com pdfplumber
    with pdfplumber.open(caminho_pdf) as pdf:

        # Abre o PDF com PyPDF2
        leitor_pdf = PdfReader(caminho_pdf)

        for i, pagina in enumerate(pdf.pages):
            try:
                # Extrai o texto da página atual
                texto_pdf = pagina.extract_text()

                # Divide o texto para extrair o número do auto
                corte_auto_1 = re.split(" / ", texto_pdf)
                corte_auto_2 = re.split(" ", corte_auto_1[1])
                auto = corte_auto_2[0]

                # Divide o texto para extrair a placa
                corte_placa_1 = re.split("Placa: ", texto_pdf)
                corte_placa_2 = re.split("\nNome:", corte_placa_1[1])
                placa = corte_placa_2[0]
                regex_valores = r'TOTAL GERAL \(R\$\): (\d{1,3}(?:\.\d{3})*,\d{2})'
                match = re.search(regex_valores, texto_pdf)
                valor = match.group(1)
                valor = f"R$ {valor}"

                # Combina a placa e o número do auto para criar o nome do arquivo
                nome_pdf = f"{placa} {auto}.pdf"

                # Expressão regular para verificar o formato de data (dd/mm/aaaa)
                padrao_data = r'\b\d{2}/\d{2}/\d{4}\b'

                if re.search(padrao_data, nome_pdf):
                    print("O nome do arquivo contém uma data.")

                    corte_auto_1 = re.split("CAMPO AUTOGESTÃO CAMPO ", texto_pdf)
                    corte_auto_2 = re.split(" ", corte_auto_1[1])
                    auto = corte_auto_2[1].replace("\nTOTAL", "")

                    # Combina a placa e o número do auto para criar o nome do arquivo
                    nome_pdf = f"{placa} {auto}.pdf"

                if ('/') in nome_pdf:
                    # Remove tudo após a barra
                    resultado = nome_pdf.split('/')[0]
                    nome_pdf = f"{resultado}.pdf"

                # Caminho completo do arquivo de saída
                caminho_pdf_saida = os.path.join(pasta_desmebrados, nome_pdf)

                # Armazenando os resultados de cada extração
                resultados_lista.append({

                    "placa": placa,
                    "auto": auto,
                    "valor": valor
                })

                # Verifica se o arquivo de saída já existe
                if not os.path.exists(caminho_pdf_saida):
                    # Criar um novo PDF para a página atual
                    escritor_pdf = PdfWriter()
                    escritor_pdf.add_page(leitor_pdf.pages[i])  # Adiciona a página correta

                    # Salva o PDF no caminho de saída
                    with open(caminho_pdf_saida, "wb") as arquivo_pdf_saida:
                        escritor_pdf.write(arquivo_pdf_saida)
                    print(f"Página {i + 1} salva no caminho: {caminho_pdf_saida}")
                else:
                    print(f"Arquivo já existe: {caminho_pdf_saida}")

            except Exception as e:
                # Salva o erro no arquivo de erros
                with open(caminho_erros, "a") as arquivo_erros:
                    arquivo_erros.write(f"Erro na página {i + 1} \n")
                print(f"Erro na página {i + 1}. Detalhes registrados em: {caminho_erros}")

#Criar um novo excel


# Criar um novo arquivo Excel
excel = Workbook()
instanciaExcel = excel.active
instanciaExcel.title = 'Dados'

# Cabeçalhos
instanciaExcel['A1'] = 'PLACA'
instanciaExcel['B1'] = 'AUTO'
instanciaExcel['C1'] = 'VALOR'

for resultado in resultados_lista:
    instanciaExcel.append([
        resultado['placa'],
        resultado['auto'],
        resultado['valor']
    ])

# Obter a data e hora atual
data_hora_atual = datetime.now()

# Formatar a data e hora para uso no nome do arquivo
data_hora_formatada = data_hora_atual.strftime("%Y-%m-%d_%H-%M-%S")

# Criar o nome do arquivo Excel
nome_excel = os.path.join(caminho_entrada, f"arquivo_{data_hora_formatada}.xlsx")

excel.save(nome_excel )

print("Processo concluído!")
