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
    dataVenc = data_encontrada[indice_escolhido_data] if len(data_encontrada) > indice_escolhido_data else None

    # Selecionar um boleto específico pelo índice
    indice_escolhido_boleto = 5
    boleto_escolhido = boletos[indice_escolhido_boleto] if len(boletos) > indice_escolhido_boleto else None

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
        boleto_escolhido, prefeituras_encontradas, ultimo_valor, data_vencimento = encontrar_boletos(texto_extraido)

        chave = boleto_escolhido if boleto_escolhido else f"pagina_{numero_pagina}"

        resultados[chave] = {
            "pagina": numero_pagina,
            "boleto": boleto_escolhido,
            "prefeitura": ", ".join(prefeituras_encontradas) if prefeituras_encontradas else "Nenhuma",
            "valor": ultimo_valor if ultimo_valor else "Nenhum",
            "Data Vencimento": data_vencimento if data_vencimento else "Nenhuma"
        }

    return resultados

# Função que processa os boletos de acordo com as diferentes condições
def processar_boletos_adicionais(texto_lista):
    for cont in range(0, len(texto_lista)):
        if "Estado do Espírito Santo - Departamento Estadual De Transito" in texto_lista[cont]:
            corteValorES1 = re.split("Pagar até:", texto_lista[cont])
            corteValorES2 = re.split(r"R\$", corteValorES1[0])
            valorMultaES = corteValorES2[1].replace("\n", "").strip()

            regexBoletoES = r"\d{11}-\d"
            valorBoletoESLista = re.findall(regexBoletoES, texto_lista[cont])
            valorBoletoESLista = list(dict.fromkeys(valorBoletoESLista))
            codigoBarrasES = ''.join(valorBoletoESLista).replace("-", "").strip()

            corteDataVencimento1 = re.split("Placa", texto_lista[cont])
            corteDataVencimento2 = re.split("Data de Vencimento", corteDataVencimento1[0])
            regexDataVencimento = r"\d{2}/\d{2}/\d{4}"
            filtrarDataVencimento = re.findall(regexDataVencimento, corteDataVencimento2[1])
            dataVencimentoES = ''.join(filtrarDataVencimento)
            print(f"ES - Código: {codigoBarrasES}, Valor: {valorMultaES}, Vencimento: {dataVencimentoES}")
        
        if "DNIT" in texto_lista[cont]:
            corteDataVencimentoDNIT1 = re.split("PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO", texto_lista[cont])
            corteDataVencimentoDNIT2 = re.split("Nome do Beneficiário/CPF/CNPJ", corteDataVencimentoDNIT1[1])
            dataVencimentoDNIT = corteDataVencimentoDNIT2[0].strip()

            corteValorDNIT1 = re.split("conforme Artigo 30 da Resolução Nº 619/16 do CONTRAN.", texto_lista[cont])
            corteValorDNIT2 = re.split("3. Até", corteValorDNIT1[1])
            valorDNIT = corteValorDNIT2[0].strip()

            corteDNITCodigo1 = re.split("Recibo do Pagador", texto_lista[cont])
            corteDNITCodigo2 = re.split("Nosso-Número",corteDNITCodigo1[1])
            codigoBarrasDNIT = corteDNITCodigo2[0].replace('\n', '').replace(' ', '').replace('-', '').replace('.', '')
            print(f"DNIT - Código: {codigoBarrasDNIT}, Valor: {valorDNIT}, Vencimento: {dataVencimentoDNIT}")

        if "DETRAN - MG" in texto_lista[cont]:
            corteValorDetran1 = re.split("Placa Marca/Modelo Data de Emissão Data de Vencimento VALOR", texto_lista[cont])
            corteValorDetran2 = re.split("Agente Data da Ocorrência Hora Local da Ocorrência Ident. Infrator", corteValorDetran1[2])
            regexValorDetran = r'R\$[\s]*[\d,.]+'
            valorDetranLista = re.findall(regexValorDetran, corteValorDetran2[0])
            valorDetran = str(valorDetranLista[0]).replace("R$", "").strip()

            regex_datas = r'\d{2}/\d{2}/\d{4}'
            datasDetranLista = re.findall(regex_datas, corteValorDetran2[0])
            dataDetran = datasDetranLista[1]

            corteDetranCodigo1 = re.split("DISQUE DENÚNCIA - 181 Guia Banco", texto_lista[cont])
            corteDetranCodigo2 = re.split("(AUTENTICAÇÃO MECÂNICA)", corteDetranCodigo1[1])
            codigoBarrasDetranMG = corteDetranCodigo2[0].replace('\n', '').replace(' ', '').replace('-', '').replace('.', '').replace('(', '')
            print(f"DETRAN - MG - Código: {codigoBarrasDetranMG}, Valor: {valorDetran}, Vencimento: {dataDetran}")

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

