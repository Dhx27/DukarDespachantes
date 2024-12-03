import os
import re
import pdfplumber

# Função para encontrar boletos no texto


def encontrar_boletos(texto):
    
    padroes_prefeitura = [
        r'(?i)\bPREFEITURA DO MUNICÍPIO DE SÃO PAULO\b',
        r'(?i)\bDETRAN - MG\b',
        r'(?i)\bPREFEITURA DE [\w\s]+\b',
        r'(?i)\bPREFEITURA DO MUNICIPIO DE BARUERI\b',
        r'(?i)\bPREFEITURA DO MUNICÍPIO DE CAJAMAR\b',
        r'(?i)\bPOLÍCIA RODOVIÁRIA FEDERAL - PRF\b'
    ]
    
    padroes_boletos = [
        r'\b\d{11}-\d\b',
        r'\b\d{5}\.\d{5}\s?\d{5}\.\d{6}\s?\d{5}\.\d{6}\s?\d\s?\d{14}\b',
        r'\b\d{44}\b',
        r'\b\d{47}\b',
        r'\d{47}',
        r'\b\d{2}\.\d{3}\.\d{3}\.\d{2}\.\d{3}\.\d{3}\b',
        r'\b\d{5}\s?\d{5}\s?\d{5}\s?\d{6}\s?\d{5}\s?\d{6}\b',
        r'\b\d{5}[-\s]?\d{5}[-\s]?\d{5}[-\s]?\d{5}\b',
        r'\b\d{11}\s\d\b'
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
    prefeituras = []
    valores_encontrados = []
    data_encontrada = []  # Armazena todas as datas encontradas

    for padrao in padroes_boletos:
        encontrados = re.findall(padrao, texto)
        for boleto in encontrados:
            boleto_limpo = re.sub(r'[-.\s,]', '', boleto)
            if boleto_limpo not in boletos:
                boletos.append(boleto_limpo)

    for padrao in padroes_prefeitura:
        match = re.search(padrao, texto)
        if match:
            prefeituras.append(match.group(0))

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
    boleto_escolhido = boletos[indice_escolhido_boleto] if len(
        boletos) > indice_escolhido_boleto else None
    # Apagar depois

    ultimo_valor = valores_encontrados[-1] if valores_encontrados else None

    return boleto_escolhido, prefeituras, ultimo_valor, dataVenc

# Função para extrair o texto de todas as páginas de um PDF


def extrair_texto_pdf(caminho_pdf):
    paginas = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for i, pagina in enumerate(pdf.pages):
            texto = pagina.extract_text()
            if texto:
                paginas.append((i + 1, texto))
    return paginas

# Função para processar boletos em um único PDF com várias páginas


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

        chave = boleto_escolhido if boleto_escolhido else f"pagina_{
            numero_pagina}"

        resultados[chave] = {
            "pagina": numero_pagina,
            "boleto": boleto_escolhido,
            "prefeitura": ", ".join(prefeituras_encontradas) if prefeituras_encontradas else "Nenhuma",
            "valor": ultimo_valor if ultimo_valor else "Nenhum",
            "Data Vencimento": data_vencimento if data_vencimento else "Nenhuma"
        }

    return resultados


# Exemplo de uso
caminho_pdf = r'M:\TI\ROBOS\ROBOS_EM_DEV\Automação Python\Boleto Py\Boletos\PDFsam_merge.pdf'
resultados_boletos = processar_boletos_em_pdf(caminho_pdf)

# Exibir resultados
for boleto, dados in resultados_boletos.items():
    print(f"\nBoleto {boleto}:")
    print(f"  Página: {dados['pagina']}")
    print(f"  Prefeitura: {dados['prefeitura']}")
    print(f"  Valor: {dados['valor']}")
    print(f"  Data de Vencimento: {dados['Data Vencimento']}")
