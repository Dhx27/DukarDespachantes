import os
import pdfplumber
import re


def extrair_informacoes_boleto(texto):
    padroes_prefeitura = [
        r'(?i)\bPREFEITURA DO MUNICÍPIO DE SÃO PAULO\b',
        r'(?i)\bDETRAN - MG\b',
        r'(?i)\bPREFEITURA DE [\w\s]+\b',
        r'(?i)\bPREFEITURA DO MUNICIPIO DE BARUERI\b',
        r'(?i)\bPREFEITURA DO MUNICÍPIO DE CAJAMAR\b',
        r'(?i)\bPOLÍCIA RODOVIÁRIA FEDERAL - PRF\b'
    ]
    prefeituras_encontradas = []

    for padrao in padroes_prefeitura:
        match = re.search(padrao, texto)
        if match:
            prefeituras_encontradas.append(match.group(0))

    if "POLÍCIA RODOVIÁRIA FEDERAL - PRF" in prefeituras_encontradas:
        print("Polícia Rodoviária Federal encontrada no boleto.")

    return prefeituras_encontradas


def extrair_texto_pdf(caminho_pdf):
    paginas = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for i, pagina in enumerate(pdf.pages):
            texto = pagina.extract_text()
            if texto:
                paginas.append((i + 1, texto))
    return paginas


def processar_boletos_em_pdf(caminho_pdf):
    resultados = {}
    paginas = extrair_texto_pdf(caminho_pdf)
    if not paginas:
        print("Nenhuma página foi processada corretamente.")
        return resultados

    for numero_pagina, texto_extraido in paginas:
        print(f"Processando página: {numero_pagina}")
        prefeituras_encontradas = extrair_informacoes_boleto(texto_extraido)

        if prefeituras_encontradas:
            resultados[numero_pagina] = {
                "pagina": numero_pagina,
                # Exemplo de prefeitura encontrada
                "prefeitura": prefeituras_encontradas[0],
                "valor": "valor_exemplo",  # Substituir pela extração real do valor
                "Data Vencimento": "data_exemplo"  # Substituir pela extração real da data
            }
    return resultados


# Exemplo de uso
caminho_pdf = r'C:\Users\stefany\Desktop\Boleto Py\Boletos\R795374817.pdf'
resultados_boletos = processar_boletos_em_pdf(caminho_pdf)

# Exibir resultados
for boleto, dados in resultados_boletos.items():
    print(f"\nBoleto {boleto}:")
    print(f"  Página: {dados['pagina']}")
    print(f"  Prefeitura: {dados['prefeitura']}")
    print(f"  Valor: {dados['valor']}")
    print(f"  Data de Vencimento: {dados['Data Vencimento']}")
else:
    if not resultados_boletos:
        print("Prefeitura não encontrada")
