import pdfplumber
from PyPDF2 import PdfReader, PdfWriter
import re
import os
from openpyxl import Workbook
from datetime import datetime

caminho_entrada = r"M:\PAGAMENTOS\SEMINOVOS\1. CP'S PARA PAGAMENTOS\DIOGO\CP 146492"

lista_pdf = [f for f in os.listdir(caminho_entrada) if f.endswith('.pdf')]

pasta_desmebrados = r"M:\PAGAMENTOS\SEMINOVOS\1. CP'S PARA PAGAMENTOS\DIOGO\CP 146492"

#instancia a lista dos resultados dos boletos
resultados_lista = []

for pdf in lista_pdf:

    nome_salva_pdf = pdf.replace(".pdf", "")

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

                # Divide o texto para extrair a placa
                regex = r"Placa:\s*([A-Z]{3}[0-9][0-9A-Z][0-9]{2})"
                
                match = re.search(regex, texto_pdf)
                placa = match.group(1) 
                    
                nome_pdf = f"{placa}.pdf"

                # Caminho completo do arquivo de saída
                caminho_pdf_saida = os.path.join(pasta_desmebrados, nome_pdf)
                
                resultados_lista.append({

                    "placa": placa
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

