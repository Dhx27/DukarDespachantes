import pdfplumber
import re
import os
from datetime import datetime, timedelta
from openpyxl import Workbook
from decimal import Decimal, ROUND_HALF_UP


codCliente = input("Insira o codigo do cliente: ")

# Função para adicionar valores ao Excel
def adicionar_ao_excel(resultados, caminho_excel):
    excel = Workbook()
    instanciaExcel = excel.active
    instanciaExcel.title = 'CNAB240'

    # Cabeçalhos
    instanciaExcel['A1'] = 'Codigo de Barras'
    instanciaExcel['B1'] = 'Orgao Autuador'
    instanciaExcel['C1'] = 'Data Vencimento'
    instanciaExcel['D1'] = 'Valor a Pagar'
    instanciaExcel['E1'] = 'Valor Desconto'
    instanciaExcel['F1'] = 'AIT'  # Corrigido para F1
    instanciaExcel['G1'] = 'Codigo Cliente'
    instanciaExcel['H1'] = 'CNPJ'

    # Adiciona os resultados
    for resultado in resultados:
        instanciaExcel.append([
            resultado['boleto'],
            resultado['nome_beneficiario'],
            resultado['data_vencimento'],
            resultado['valor'],
            resultado['desconto'],
            resultado['ait'],
            resultado['codigo'],
            resultado['cnpj']
        ])

    # Salva o arquivo Excel
    excel.save(caminho_excel)
    print(f"Dados salvos em {caminho_excel}")


caminho_pdf = r"M:\TI\ROBOS\ROBOS_EM_DEV\Automação Python\Boleto Py\data\boletosConcatenados\PDFsam_merge.pdf"
texto_lista = []
resultados = []  # Lista para armazenar os resultados



# Extração do texto do PDF
with pdfplumber.open(caminho_pdf) as pdf:
    for numero_pagina, pagina in enumerate(pdf.pages, start=1):
        texto_pagina = pagina.extract_text()
        if texto_pagina:
            texto_lista.append(texto_pagina)
        else:
            texto_lista.append(f"BOLETO INDISPONIVEL, FAVOR VERIFICAR NA PAGIN: {numero_pagina}" )

# Função para encontrar boletos, valores e data de vencimento


def encontrar_boletos(texto):

    texto = re.sub(r'\d{2}/\d{2}/\d{4}\d{2}:\d{2}:\d{2}', '', texto)

    padroes_boletos = [
        # Captura números de 11 dígitos seguidos de um hífem e 1 dígito
        r'\b\d{11}-\d\b',
        # Captura o padrão de números com pontos e espaços
        r'\b\d{5}\.\d{5}\s?\d{5}\.\d{6}\s?\d{5}\.\d{6}\s?\d\s?\d{14}\b',
        # Captura apenas números com 44 ou 47 dígitos
        r'\b\d{44}\b|\b\d{47}\b',
        r'\d{47}',  # Captura qualquer sequência de exatamente 47 dígitos
        # Captura CPF ou CNPJ no formato padrão
        r'\b\d{2}\.\d{3}\.\d{3}\.\d{2}\.\d{3}\.\d{3}\b',
        # Captura um número longo com espaços
        r'\b\d{5}\s?\d{5}\s?\d{5}\s?\d{6}\s?\d{5}\s?\d{6}\b',
        # Captura números com hífens ou espaços
        r'\b\d{5}[-\s]?\d{5}[-\s]?\d{5}[-\s]?\d{5}\b',
        # Captura números de 11 dígitos seguidos de um espaço e 1 dígito
        r'\b\d{11}\s\d\b'
        # Novo padrão adicionado: captura números com 14 dígitos (sem separadores)
        r'\b\d{14}\b'
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
        encontrados = re.findall(padrao, texto)

        for boleto in encontrados:
            # Remove espaços, vírgulas e hífens
            boleto_limpo = re.sub(r'[-\s,]', '', boleto)
            if boleto_limpo not in boletos:
                boletos.append(boleto_limpo)

    # Agora, você pode juntar todos os boletos limpos em uma única string
    boleto_unico = ''.join(boletos)

    # Agora, você pode juntar todos os boletos limpos em uma única string
    if boletos:
        boleto_unico = ''.join(boletos)  # Junta todos os boletos na string
    else:
        boleto_unico = None  # Se não houver boletos, define como None

    for padrao in padroes_valor:
        valores_encontrados.extend(re.findall(padrao, texto))

    for padrao in padroes_DataVenc:
        data_encontrada.extend(re.findall(padrao, texto))

    # Verifica se existem boletos e valores encontrados
    boleto_escolhido = boleto_unico if boletos else None
    ultimo_valor = valores_encontrados[-1] if valores_encontrados else None
    dataVenc = data_encontrada[0] if data_encontrada else None

    return boleto_escolhido, ultimo_valor, dataVenc


# Processa o texto extraído e armazena os resultados em uma lista
for cont in range(len(texto_lista)):


    if "Estado do Espírito Santo - Departamento Estadual De Transito" in texto_lista[cont]:

        corteValorES1 = re.split("Pagar até:", texto_lista[cont])
        corteValorES2 = re.split(r"R\$", corteValorES1[0])
        valorMultaES = corteValorES2[1].replace("\n", "").strip()

        corteDescontoES1 = re.split("Multa UF:", texto_lista[cont])
        corteDescontoES2 = re.split("Total a Pagar", corteDescontoES1[1])
        corteDescontoES3 = re.split(" ", corteDescontoES2[0])
        valorDescontoES = corteDescontoES3[5]

        regexBoletoES = r"\d{11}-\d"
        valorBoletoESLista = re.findall(regexBoletoES, texto_lista[cont])
        valorBoletoESLista = list(dict.fromkeys(valorBoletoESLista))
        codigoBarrasES = ''.join(valorBoletoESLista).replace("-", "").strip()

        corteDataVencimento1 = re.split("Placa", texto_lista[cont])
        corteDataVencimento2 = re.split(
            "Data de Vencimento", corteDataVencimento1[0])
        regexDataVencimento = r"\d{2}/\d{2}/\d{4}"
        filtrarDataVencimento = re.findall(
            regexDataVencimento, corteDataVencimento2[1])
        dataVencimentoES = ''.join(filtrarDataVencimento)
        
        corteAutoES1 = re.split("Multa UF:",  texto_lista[cont])
        corteAutoES2 = re.split("/", corteAutoES1[1])
        corteAutoES3 = re.split("-", corteAutoES2[0])
        autoES = corteAutoES3[2]
        
        nomeBeneficiario = "DETRAN-ES"

        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': codigoBarrasES,
            'valor': valorMultaES,
            'desconto' : valorDescontoES,
            'data_vencimento': dataVencimentoES,
            'nome_beneficiario': nomeBeneficiario,
            'ait': autoES,
            'codigo': codCliente,
            'cnpj': "28162105000166"
            })

    elif "DNIT" in texto_lista[cont]:


        regex_cnpj = r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}'

        lista_cnpjDNIT = re.findall(regex_cnpj, texto_lista[cont])

        cnpj_dnit = lista_cnpjDNIT[0].replace(".", "").replace("/", "").replace("-", "")

        corteDataVencimentoDNIT1 = re.split(
            "PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO", texto_lista[cont])
        corteDataVencimentoDNIT2 = re.split(
            "Nome do Beneficiário/CPF/CNPJ", corteDataVencimentoDNIT1[1])
        dataVencimentoDNIT = corteDataVencimentoDNIT2[0].strip()

        corteValorDNIT1 = re.split(
            "conforme Artigo 30 da Resolução Nº 619/16 do CONTRAN.", texto_lista[cont])
        corteValorDNIT2 = re.split("3. Até", corteValorDNIT1[1])
        valorDNIT = corteValorDNIT2[0].strip()

        corteDescontoDNIT1 = re.split("Desconto/Abatimento", texto_lista[cont])
        corteDescontoDNIT2 = re.split("1.", corteDescontoDNIT1[1].replace('\n', ''))
        valorDescontoDNIT = corteDescontoDNIT2[0]

        corteDNITCodigo1 = re.split("Recibo do Pagador", texto_lista[cont])
        corteDNITCodigo2 = re.split("Nosso-Número", corteDNITCodigo1[1])
        codigoBarrasDNIT = corteDNITCodigo2[0].replace(
            '\n', '').replace(' ', '').replace('-', '').replace('.', '')
        
        corteAutoDNIT1 = re.split("Número do Auto Placa Marca / Modelo", texto_lista[cont])
        corteAutoDNIT2 = re.split("/", corteAutoDNIT1[1])
        corteAutoDNIT3 = re.split(" ", corteAutoDNIT2[0])
        AutoDNIT = corteAutoDNIT3[0].replace("\n", "")
        
        nomeBeneficiario = "DNIT"

        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': codigoBarrasDNIT,
            'valor': valorDNIT,
            'desconto' : valorDescontoDNIT,
            'data_vencimento': dataVencimentoDNIT,
            'nome_beneficiario': nomeBeneficiario,
            'ait': AutoDNIT,
            'codigo': codCliente,
            'cnpj': cnpj_dnit
            
        })

    elif "DETRAN - MG" in texto_lista[cont]:

        corteValorDetran1 = re.split(
            "Placa Marca/Modelo Data de Emissão Data de Vencimento VALOR", texto_lista[cont])
        corteValorDetran2 = re.split(
            "Agente Data da Ocorrência Hora Local da Ocorrência Ident. Infrator", corteValorDetran1[2])
        regexValorDetran = r'R\$[\s]*[\d,.]+'
        valorDetranLista = re.findall(regexValorDetran, corteValorDetran2[0])
        valorDetran = str(valorDetranLista[0]).replace("R$", "").strip()

        regex_datas = r'\d{2}/\d{2}/\d{4}'
        datasDetranLista = re.findall(regex_datas, corteValorDetran2[0])
        dataDetran = datasDetranLista[1]

        corteDetranCodigo1 = re.split(
            "DISQUE DENÚNCIA - 181 Guia Banco", texto_lista[cont])
        corteDetranCodigo2 = re.split(
            "(AUTENTICAÇÃO MECÂNICA)", corteDetranCodigo1[1])
        codigoBarrasDetranMG = corteDetranCodigo2[0].replace('\n', '').replace(
            ' ', '').replace('-', '').replace('.', '').replace('(', '')
        
        AutoDetranMG = "NA"
        
        nomeBeneficiario = "DETRAN - MG"

        '''
        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': codigoBarrasDetranMG,
            'valor': valorDetran,
            'data_vencimento': dataDetran,
            'nome_beneficiario': nomeBeneficiario,
            'ait': AutoDetranMG
        })
        '''

    elif "POLÍCIA RODOVIÁRIA FEDERAL - PRF" in texto_lista[cont]:
            
        print(texto_lista[cont])
    
        corteCNPJ1 = re.split("CNPJ: ", texto_lista[cont])
        corteCNPJ2 = re.split("\nAutenticação mecânica", corteCNPJ1[1])
        cnpj_PRF = corteCNPJ2[0]

        corteValorPRF1 = re.split("Valor Cobrado:", texto_lista[cont])
        corteValorPRF2 = re.split(
            "Instruções para pagamento:", corteValorPRF1[1])
        regexValorPRF = r'R\$[\s]*[\d,.]+'
        valorPRFLista = re.findall(regexValorPRF, corteValorPRF2[1])
        valorPRF = str(valorPRFLista[0]).replace("R$", "").strip()

        corteDescontoPRF1 = re.split("Descontos/Abatimento :", texto_lista[cont])
        corteDescontoPRF2 = re.split("\n", corteDescontoPRF1[1])
        descontoPRF = corteDescontoPRF2[0].replace('R$ ', '')

        cortePRFCodigo1 = re.split("PAGAMENTO EM CHEQUE", texto_lista[cont])
        cortePRFCodigo2 = re.split("Vencimento", cortePRFCodigo1[1])
        codigoBarrasPRF = cortePRFCodigo2[0].replace(
            '\n|001-9|', '').replace('\n', '').strip()

        dataVencimentoPRF1 = re.split(
            "PAGÁVEL NA REDE BANCÁRIA ATÉ O VENCIMENTO.", texto_lista[cont])
        dataVencimentoPRF2 = re.split(
            "Nome Beneficiário: Agência/Cód. Beneficiário", dataVencimentoPRF1[1])
        dataVencimentoPRF = dataVencimentoPRF2[0].replace('\n', '')
        
        corteAutoPRF1 = re.split("Auto de Infração NIT/NAP Peso Excedente Velocidade", texto_lista[cont])
        corteAutoPRF2 = re.split("CNH Condutor CPF/CNPJ Propretário VIN", corteAutoPRF1[1])
        corteAutoPRF3 = re.split(" ", corteAutoPRF2[0])
        AutoPRF = corteAutoPRF3[0].replace("\n", "")
        
        nomeBeneficiario = "PRF"

        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': codigoBarrasPRF,
            'valor': valorPRF,
            'desconto': descontoPRF,
            'data_vencimento': dataVencimentoPRF,
            'nome_beneficiario': nomeBeneficiario,
            'ait': AutoPRF,
            'codigo': codCliente,
            'cnpj': cnpj_PRF
        })

    elif "ESTADODOPARANÁ" in texto_lista[cont]:
        
        corteAutoPARANA1 = re.split("AutodeInfração: Situação:", texto_lista[cont])
        corteAutoPARANA2 = re.split("Data/HoraInfração: ÓrgãoCompetente:", corteAutoPARANA1[1])
        corteAutoPARANA3 = re.split(" ", corteAutoPARANA2[0])
        corteAutoPARANA4 = re.split("-", corteAutoPARANA3[0])
        AutoPARANA = corteAutoPARANA4[1]
        
        boleto, valor, dataVenc = encontrar_boletos(texto_lista[cont])
        
        nomeBeneficiario = "DETRAN / PR"

        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': boleto,
            'valor': valor,
            'desconto': "0,00",
            'data_vencimento': dataVenc,
            'nome_beneficiario': nomeBeneficiario,
            'ait': AutoPARANA,
            'codigo': codCliente,
            'cnpj': "78206513000140"
        })
        
    elif "SENATRAN SNE" in texto_lista[cont]:
        
        print(texto_lista[cont])
        
        corteAutoSENATRAN1 = re.split("Auto de Infração: ", texto_lista[cont])
        corteAutoSENATRAN2 = re.split(" Valor:", corteAutoSENATRAN1[1])
        AutoSenatran = corteAutoSENATRAN2[0]

        corteDescontoSENATRAN1 = re.split("Valor:", texto_lista[cont])
        corteDescontoSENATRAN2 =  re.split("Código da Infração:", corteDescontoSENATRAN1[1])
        valorTotalSENATRAN = corteDescontoSENATRAN2[0].replace(' R$ ', '').replace('\n', '')

        boleto, valor, dataVenc = encontrar_boletos(texto_lista[cont])
        
        valorLimpo = valor.replace(',', '.').strip()
        desconto = float(valorLimpo)

        valorTotalLimpo = valorTotalSENATRAN.replace(',', '.').strip()
        total = float(valorTotalLimpo)

        valorDescontoSENATRAN = total - desconto
        valorDescontoSENATRAN = round(valorDescontoSENATRAN, 2)  # Arredonda para 2 casas decimais

        valorDescontoSENATRAN = str(valorDescontoSENATRAN).replace('.', ',')

        nomeBeneficiario = "SENATRAN SNE"

        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': boleto,
            'valor': valor,
            'desconto': valorDescontoSENATRAN,
            'data_vencimento': dataVenc,
            'nome_beneficiario': nomeBeneficiario,
            'ait': AutoSenatran,
            'codigo': codCliente,
            'cnpj': 0
        })
        
#resultados.sort(key=lambda x: x['boleto'], reverse=True)


# Exibindo os resultados
for resultado in resultados:
    print(resultado)
    
caminho_excel = r'C:\Users\diogo.lana\Desktop\Diogo\TESTE2.xlsx'


# Adicionar os dados ao Excel
adicionar_ao_excel(resultados, caminho_excel)

# Calcular a soma total de todos os valores
total_valores = sum(float(resultado['valor'].replace(
    ',', '.')) for resultado in resultados if resultado['valor'])
print(f"Soma total dos valores: R$ {total_valores:.2f}")

