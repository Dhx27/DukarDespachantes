import pdfplumber
import re
import os
from datetime import datetime, timedelta

caminho_pdf = r"M:\TI\ROBOS\ROBOS_EM_DEV\Automação Python\Boleto Py\Boletos\PDFsam_merge.pdf"
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
        
        nomeBeneficiario = "DETRAN-ES"

        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': codigoBarrasES,
            'valor': valorMultaES,
            'data_vencimento': dataVencimentoES,
            'nome_beneficiario': nomeBeneficiario
        })

    elif "DNIT" in texto_lista[cont]:
        corteDataVencimentoDNIT1 = re.split(
            "PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO", texto_lista[cont])
        corteDataVencimentoDNIT2 = re.split(
            "Nome do Beneficiário/CPF/CNPJ", corteDataVencimentoDNIT1[1])
        dataVencimentoDNIT = corteDataVencimentoDNIT2[0].strip()

        corteValorDNIT1 = re.split(
            "conforme Artigo 30 da Resolução Nº 619/16 do CONTRAN.", texto_lista[cont])
        corteValorDNIT2 = re.split("3. Até", corteValorDNIT1[1])
        valorDNIT = corteValorDNIT2[0].strip()

        corteDNITCodigo1 = re.split("Recibo do Pagador", texto_lista[cont])
        corteDNITCodigo2 = re.split("Nosso-Número", corteDNITCodigo1[1])
        codigoBarrasDNIT = corteDNITCodigo2[0].replace(
            '\n', '').replace(' ', '').replace('-', '').replace('.', '')
        
        nomeBeneficiario = "DNIT"

        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': codigoBarrasDNIT,
            'valor': valorDNIT,
            'data_vencimento': dataVencimentoDNIT,
            'nome_beneficiario': nomeBeneficiario
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

        nomeBeneficiario = "DETRAN - MG"

        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': codigoBarrasDetranMG,
            'valor': valorDetran,
            'data_vencimento': dataDetran,
            'nome_beneficiario': nomeBeneficiario
        })

    elif "POLÍCIA RODOVIÁRIA FEDERAL - PRF" in texto_lista[cont]:

        corteValorPRF1 = re.split("Valor Cobrado:", texto_lista[cont])
        corteValorPRF2 = re.split(
            "Instruções para pagamento:", corteValorPRF1[1])
        regexValorPRF = r'R\$[\s]*[\d,.]+'
        valorPRFLista = re.findall(regexValorPRF, corteValorPRF2[1])
        valorPRF = str(valorPRFLista[0]).replace("R$", "").strip()

        cortePRFCodigo1 = re.split("PAGAMENTO EM CHEQUE", texto_lista[cont])
        cortePRFCodigo2 = re.split("Vencimento", cortePRFCodigo1[1])
        codigoBarrasPRF = cortePRFCodigo2[0].replace(
            '\n|001-9|', '').replace('\n', '').strip()

        dataVencimentoPRF1 = re.split(
            "PAGÁVEL NA REDE BANCÁRIA ATÉ O VENCIMENTO.", texto_lista[cont])
        dataVencimentoPRF2 = re.split(
            "Nome Beneficiário: Agência/Cód. Beneficiário", dataVencimentoPRF1[1])
        dataVencimentoPRF = dataVencimentoPRF2[0].replace('\n', '')

        nomeBeneficiario = "PRF"

        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': codigoBarrasPRF,
            'valor': valorPRF,
            'data_vencimento': dataVencimentoPRF,
            'nome_beneficiario': nomeBeneficiario
        })

    elif "ESTADODOPARANÁ" in texto_lista[cont]:

        boleto, valor, dataVenc = encontrar_boletos(texto_lista[cont])
        
        nomeBeneficiario = "DETRAN / PR"

        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': boleto,
            'valor': valor,
            'data_vencimento': dataVenc,
            'nome_beneficiario': nomeBeneficiario
        })
        
    elif "SENATRAN SNE" in texto_lista[cont]:
        
        boleto, valor, dataVenc = encontrar_boletos(texto_lista[cont])
        
        nomeBeneficiario = "SENATRAN SNE"

        # Armazenando os resultados de cada extração
        resultados.append({
            'boleto': boleto,
            'valor': valor,
            'data_vencimento': dataVenc,
            'nome_beneficiario': nomeBeneficiario
        })
        

# Exibindo os resultados
for resultado in resultados:
    print(resultado)

# Calcular a soma total de todos os valores
total_valores = sum(float(resultado['valor'].replace(
    ',', '.')) for resultado in resultados if resultado['valor'])
print(f"Soma total dos valores: R$ {total_valores:.2f}")

# Obter a data e hora atual
dia_atual = datetime.now()
# Converter para string no formato desejado (exemplo: "HHMMSS")
hora_formatada = dia_atual.strftime("%H%M%S")

dia_atual = datetime.now()
# Converter para string no formato desejado (exemplo: "DDMMYYYY")
data_formatada = dia_atual.strftime("%d%m%Y")

# Definindo o nome do arquivo e o arquivo de controle do lote
nome_arquivo = "Pagamento_bb.txt"
arquivo_controle = "controle_lote.txt"

# Informações para o Header do arquivo

header_arquivo = [
    "001",  # Código do Banco na Compensação
    "0000",  # Lote de Serviço
    "0",  # Tipo de Serviço
    "         ",  # Uso Exclusivo FEBRABAN/CNAB
    "2",  # Tipo de Inscrição da Empresa
    "13022408000106",  # Número de Inscrição da Empresa
    "000133624",  # Código do Convênio no Banco
    # "00000000000000133624",  # Nº do Convênio
    "0126",  # Código
    "     ",  # Uso Reservado do Banco
    "  ",  # Arquivo de teste
    "38830",  # Agência Mantenedora da Conta
    "X",  # Dígito Verificador da Agência
    "000000123597",  # Número da Conta Corrente (completar zeros a esquerda)
    "X",  # Dígito Verificador da Conta
    "0",  # Dígito Verificador da Ag/Conta
    "DUT KAR SERVIÇOS LTDA EPP     ",  # Nome da Empresa
    "BANCO DO BRASIL               ",  # Nome do Banco
    "          ",  # Uso Exclusivo FEBRABAN/CNAB
    "1",  # Código Remessa / Retorno
    f"{data_formatada}",  # Data de Geração do Arquivo
    f"{hora_formatada}",  # Hora de Geração do Arquivo
    "000001",  # Número Sequencial do Arquivo
    "240",  # Nº Versão do Layout do Arquivo
    "00000",  # Densidade de Gravação do Arquivo
    "                   ",  # Para Uso Reservado do Banco
    "                    ",  # Para Uso Reservado da Empresa
    "           ",  # Para Uso Exclusivo FEBRABAN CNAB
    "   ",  # Identificação cobrança sem papel
    "000",  # Uso exclusivo das VANS
    "00",  # Tipo de Serviço
    "0000000000",  # Códigos de Ocorrências
]

# Informações para o Header de Lote
header_lote = [
    "001",  # Código do Banco
    "0001",  # Lote de serviço (formatado com 4 dígitos)
    "1",  # Tipo de registro
    "C",  # Tipo da Operação
    "98",  # Tipo do Serviço
    "31",  # Forma de Lançamento
    "240",  # N° da Versão do Layout do Lote
    " ",  # Uso Exclusivo da CNAB
    "2",  # Tipo de Inscrição da Empresa
    "13022408000106",  # Número de Inscrição da Empresa
    # "????????????????????",  # Código do Convênio no Banco
    "000133624",  # Nº do Convênio
    "0126",  # Código
    "     ",  # Uso Reservado do Banco
    "  ",  # Arquivo de teste
    "38830",  # Agência Mantenedora da Conta
    "X",  # Dígito Verificador da Agência
    "000001235974",  # Número da Conta Corrente
    "X",  # Dígito Verificador da Conta
    "0",  # Dígito Verificador da Ag/Conta
    "DUT KAR SERVIÇOS LTDA EPP     ",  # Nome da Empresa
    "                                        ",  # Mensagem
    "Dr. Gordiano                  ",  # Nome da Rua, Av, Pça, Etc
    "00164",  # Número do Local
    "SALA 02        ",  # Casa, Apto, Sala, Etc
    "BELO HORIZONTE      ",  # Cidade
    "30411",  # CEP
    "080",  # Complemento do CEP
    "MG",  # Sigla do Estado
    "        ",  # Uso Exclusivo da CNAB
    "0000000000",  # Código das Ocorrências p/ Retorno
]

# Lista para armazenar os registros do seguimento J
registros_seguimento_J = []

quantBoleto = 0

for resultado in resultados:
    boleto = resultado['boleto']
    valor = resultado['valor']
    beneficiario = resultado['nome_beneficiario']

    beneficiario_formatado = beneficiario.ljust(30)
    
    if boleto is not None and valor is not None:
        data_vencimento = resultado['data_vencimento'].replace('/', '')
        
        quantBoleto += 1

        valorLista = re.split(",", valor)

        # Complementa com zeros à esquerda até que tenha 13 caracteres
        if len(valorLista[0]) <= 13:
            valorLista[0] = valorLista[0].zfill(13)

        # Combina os valores com a vírgula como separador
        valorFinal = ''.join([str(valorLista[0]), str(valorLista[1])])

        # Formata o registro do Seguimento J
        seguimento_J1 = [
            "001",  # Código do Banco
            "0001",  # Lote de serviço (mesmo número de lote)
            "3",  # Tipo de registro
            # Número do registro (incremental)
            f"{len(registros_seguimento_J) + 1:05d}",
            "J",  # Código de Segmento no Reg. Detalhe
            "0",  # Tipo de movimento
            "00",  # Código da Instrução p/ Movimento
            f"{boleto}",  # Código de Barras
            # Nome do Beneficiário vamos ter que colocar de onde é o boleto
            f"{beneficiario_formatado}",
            f"{data_vencimento}",  # Data Nominal do Vencimento
            f"{valorFinal}",  # Valor do Título
            "000000000000000",  # Valor do Desconto + Abatimento
            f"{valorFinal}",  # Valor da Mora + Multa
            f"{data_formatada}",  # Data do Pagamento
            f"{valorFinal}",  # Valor do Pagamento
            "000000000000000",  # Quantidade da Moeda
            "pagamento de boletos",  # N° Docto Atribuído pela Empresa
            "                    ",  # N° Docto Atribuído pelo Banco
            "09",  # Código de Moeda
            "      ",  # CNAB
            "0000000000",  # Código das Ocorrências p/ Retorno
        ]

        # Adiciona o registro à lista
        registros_seguimento_J.append(''.join(seguimento_J1))

# Calcula os valores necessários para o trailer
# Inclui Header e Trailer do Lote
quantidade_registros_lote = len(registros_seguimento_J) + 2
quantidade_registros_arquivo = quantidade_registros_lote + \
    2  # Inclui Header e Trailer do Arquivo

#Formata a quantidade de boletos para o padrão
quantidade_boletos_formatados = str(quantBoleto).zfill(6)    
       

# Calcula a soma total dos valores
try:
    somatoria_valores = sum(
        int(re.sub(r"[^\d]", "", resultado['valor']).zfill(15)) for resultado in resultados if resultado['valor']
    )
except Exception as e:
    print(f"Erro ao calcular somatória dos valores: {e}")
    somatoria_valores = 0

# Atualiza o Trailer do Lote
trailer_lote = [
    "001",  # Código do Banco
    "00001",  # Lote de serviço
    "5",  # Tipo de registro
    "         ",  # Uso Exclusivo CNAB
    f"{quantidade_registros_lote:06d}",  # Quantidade de Registros do Lote
    f"{somatoria_valores:018d}",  # Somatória dos Valores
    "000000000000000000",  # Somatória de Quantidade de Moedas
    "000000",  # Número Aviso Débito
    "                                                                                                                                                                     ",  # Uso Exclusivo FEBRABAN/CNAB
    "0000000000",  # Códigos das Ocorrências para Retorno
]

# Atualiza o Trailer do Arquivo
trailer_arquivo = [
    "001",  # Código do Banco na Compensação
    "9999",  # Lote de Serviço
    "9",  # Tipo de registro
    "         ",  # Excusivo CNAB
    f"{quantBoleto:6d}",  # Quantidade de boletos no lote (campo com 6 dígitos)
    "000001",  # Quantidade de Lotes do Arquivo (1 lote neste caso)
    # Quantidade de Registros do Arquivo
    f"{quantidade_registros_arquivo:06d}",
    "000000",  # Qtde. de Contas p/ Conc. (Lotes)
    " " * 81,  # Uso Exclusivo FEBRABAN / CNAB
]

# Abrindo o arquivo para escrita
with open(nome_arquivo, 'w') as arquivo:
    # Escrevendo as informações no arquivo
    arquivo.write(''.join(header_arquivo) + '\n')
    arquivo.write(''.join(header_lote) + '\n')
    for seguimento in registros_seguimento_J:
        arquivo.write(seguimento + '\n')
    arquivo.write(''.join(trailer_lote) + '\n')
    arquivo.write(''.join(trailer_arquivo) + '\n')

print(f"As informações do modelo CNAB 240 foram escritas no arquivo {
      nome_arquivo}.")
