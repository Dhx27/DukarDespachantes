import os
from datetime import datetime, timedelta



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
    "02363725000170",  # Número de Inscrição da Empresa
    "",  # Código do Convênio no Banco
    "00000000000000133624",  # Nº do Convênio
    "0126",  # Código
    "     ",  # Uso Reservado do Banco
    "TS",  # Arquivo de teste
    "?????",  # Agência Mantenedora da Conta
    "?",  # Dígito Verificador da Agência
    "????????????",  # Número da Conta Corrente (completar zeros a esquerda)
    "?",  # Dígito Verificador da Conta
    "0",  # Dígito Verificador da Ag/Conta
    "Dukar despachantes LTDA       ",  # Nome da Empresa
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
    "02363725000170",  # Número de Inscrição da Empresa
    "????????????????????",  # Código do Convênio no Banco
    "000133624",  # Nº do Convênio
    "0126",  # Código
    "     ",  # Uso Reservado do Banco
    "TS",  # Arquivo de teste
    "?????",  # Agência Mantenedora da Conta
    "?",  # Dígito Verificador da Agência
    "????????????",  # Número da Conta Corrente
    "?",  # Dígito Verificador da Conta
    "0",  # Dígito Verificador da Ag/Conta
    "DUKAR DESPACHANTES LTDA       ",  # Nome da Empresa
    "                                        ",  # Mensagem
    "Dr. Gordiano                  ",  # Nome da Rua, Av, Pça, Etc
    "164                           ",  # Número do Local
    "SALA 02        ",  # Casa, Apto, Sala, Etc
    "BELO HORIZONTE      ",  # Cidade
    "30411",  # CEP
    "080",  # Complemento do CEP
    "MG",  # Sigla do Estado
    "        ",  # Uso Exclusivo da CNAB
    "0000000000",  # Código das Ocorrências p/ Retorno
]

# Informações para o Seguimento J
seguimento_J1 = [
    "001",  # Código do Banco
    "0001",  # Lote de serviço (mesmo número de lote)
    "3",  # Tipo de registro
    "00001",  # Número do registro
    "J",  # Código de Segmento no Reg. Detalhe
    "0",  # Tipo de movimento
    "00",  # Código da Instrução p/ Movimento
    "876200000019314600002419800660095012527107074553",  # Código de Barras
    "",  # Nome do Beneficiário
    "30102024",  # Data Nominal do Vencimento
    "000000000013016",  # Valor do Título
    "000000000000000",  # Valor do Desconto + Abatimento
    "000000000000146",  # Valor da Mora + Multa
    "30102024",  # Data do Pagamento
    "000000000013146",  # Valor do Pagamento
    "000000000000000",  # Quantidade da Moeda
    "pagamento de boletos",  # N° Docto Atribuído pela Empresa
    "                    ",  # N° Docto Atribuído pelo Banco
    "09",  # Código de Moeda
    "      ",  # CNAB
    "0000000000",  # Código das Ocorrências p/ Retorno
]

# Informações para o Trailer do Lote
trailer_lote = [
    "001",  # Código do Banco
    "00001",  # Lote de serviço
    "5",  # Tipo de registro
    "         ",  # Uso Exclusivo CNAB
    "000001",  # Quantidade de Registros do Lote (formatado com 6 dígitos)
    "000000000000013146",  # Somatória dos Valores
    "000000000000000000",  # Somatória de Quantidade de Moedas
    "000000",  # Número Aviso Débito
    "                                                                                                                                                                     ",  # Uso Exclusivo FEBRABAN/CNAB
    "0000000000",  # Códigos das Ocorrências para Retorno
]

# Informações para o Trailer do Lote

trailer_arquivo = [
    "001",  # Código do Banco na Compensação
    "9999",  # Lote de Serviço
    "000001",  # Quantidade de Lotes do Arquivo
    "000002",  # Quantidade de Registros do Arquivo
    "000000",  # Qtde. de Contas p/ Conc. (Lotes)
    "                                                                                                                                                                                                             ",  # Uso Exclusivo FEBRABAN / CNAB

]

# Abrindo o arquivo para escrita
with open(nome_arquivo, 'w') as arquivo:
    # Escrevendo as informações no arquivo
    arquivo.write(''.join(header_arquivo) + '\n')
    arquivo.write(''.join(header_lote) + '\n')
    arquivo.write(''.join(seguimento_J1) + '\n')
    arquivo.write(''.join(trailer_lote) + '\n')
    arquivo.write(''.join(trailer_arquivo) + '\n'),0
    

print(f"As informações do modelo CNAB 240 foram escritas no arquivo {
      nome_arquivo}.")
