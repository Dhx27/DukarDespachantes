import pdfplumber
import re


def extrair_texto_pdf(caminho_pdf):

    paginas = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for i, pagina in enumerate(pdf.pages):
            texto = pagina.extract_text()
            if texto:
                paginas.append((i + 1, texto))
            else:
                print(f"Texto não extraído na página {i + 1}")
    return paginas


def encontrar_prefeituras(texto):

    padrao_PRF = r'(?i)\bPOLÍCIA RODOVIÁRIA FEDERAL - PRF\b'
    padrao_Detran_Mg = r'(?i)\bDETRAN - MG\b'
    padrao_Detran_ES = r'(?i)\bEstado do Espírito Santo - Departamento Estadual De Transito\b'
    padrao_Senatran = r'(?i)\bDENATRAN SNE\b'
    padrao_DNIT = r'\bDNIT\b'

    resultados = []
    for padrao in [padrao_PRF, padrao_Detran_Mg, padrao_Detran_ES, padrao_Senatran, padrao_DNIT]:
        encontradas = re.findall(padrao, texto)
        for prefeitura in encontradas:
            if prefeitura not in resultados:
                resultados.append(prefeitura)

    return resultados


def encontrar_boletos(texto):

    padroes_boletos = [
        r'(\d{11,12})(?=\s\d\s)'
    ]

    padroes_valor = [
        r'\b\d{1,3}(?:\.\d{3})?,\d{2}\b',
        r'\b\d+,\d{2}\b',
        r'R\$\s*\d+,\d{2}',
        r'(?<!\d)(R\$\s*)?\d{1,3}(?:\.\d{3})*(?:,\d{2})?(?!\d)',
        r'(?<!\d)(R\$\s*)?\d+(?:,\d{2})?(?!\d)',
        r'R\$\s*\d{1,3}(?:\.\d{3})*,\d{2}',
        r'(\d+,\d{2})'
    ]

    padroes_DataVenc = [
        r'(\d{2}/\d{2}/\d{4})',
        r'(?i)ATE\s*(\d{2}/\d{2}/\d{4})',
        r'(?i)2-VENCIMENTO\s*(\d{2}/\d{2}/\d{4})',
        r'\d{2}\.\d{2}\.\d{4}',
        r'\b\d{2}/\d{2}/\d{4}\b'
    ]

    boletos = []
    valores_encontrados = []
    data_encontrada = []

    for padrao in padroes_boletos:
        boletos = re.findall(padrao, texto)

    for padrao in padroes_valor:
        valores_encontrados.extend(re.findall(padrao, texto))

    for padrao in padroes_DataVenc:
        data_encontrada.extend(re.findall(padrao, texto))

    # Selecionar uma data específica pelo índice
    indice_escolhido_data = 4
    dataVenc = data_encontrada[indice_escolhido_data] if len(
        data_encontrada) > indice_escolhido_data else None

    # Selecionar um boleto específico pelo índice
    indice_escolhido_boleto = 5
    boletos = boletos[indice_escolhido_boleto] if len(
        boletos) > indice_escolhido_boleto else None

    ultimo_valor = valores_encontrados[-1] if valores_encontrados else None

    return boletos, ", ".join(encontrar_prefeituras(texto)), ultimo_valor, dataVenc


def processar_boletos_em_pdf(caminho_pdf):

    resultados = {}
    paginas = extrair_texto_pdf(caminho_pdf)

    if not paginas:
        print("Nenhuma página foi processada corretamente.")
        return resultados

    for numero_pagina, texto_extraido in paginas:
        print(f"Processando página: {numero_pagina}")
        boleto_escolhido, prefeituras_encontradas, ultimo_valor, data_vencimento = encontrar_boletos(
            texto_extraido)

        chave = f"pagina_{numero_pagina}"

        resultados[chave] = {
            "pagina": numero_pagina,
            "boleto": boleto_escolhido,
            "prefeitura": prefeituras_encontradas if prefeituras_encontradas else "Nenhuma",
            "valor": ultimo_valor if ultimo_valor else "Nenhum",
            "Data Vencimento": data_vencimento if data_vencimento else "Nenhuma"
        }

    return resultados


# Caminho do arquivo PDF
caminho_pdf = r'C:\Users\stefany\Desktop\Boleto Py\Boletos\FA01706219.pdf'

# Processar o PDF
resultados_boletos = processar_boletos_em_pdf(caminho_pdf)

# Exibindo os resultados formatados
for chave, dados in resultados_boletos.items():
    print(f"\n{chave}:")
    print(f"Página: {dados['pagina']}")
    print(f"Boleto: {dados['boleto']}")
    print(f"Prefeitura(s): {dados['prefeitura']}")
    print(f"Valor: {dados['valor']}")
    print(f"Data de Vencimento: {dados['Data Vencimento']}")
