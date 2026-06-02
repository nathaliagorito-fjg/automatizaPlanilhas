import pandas as pd
import unicodedata as uni
from openpyxl import load_workbook

planilhaMensal = None
planilhaMinibio = None
planilhaErgon = None
planilhaHistorico = None
diretorioHistorico = None

def normalizaTexto(texto):
    texto = str(texto).upper()
    texto = ''.join(c for c in uni.normalize('NFD', texto) if uni.category(c) != 'Mn')
    texto = texto.replace('-', '').replace('.', '')

    return texto

def criaSigla(palavra, sigla):
    global planilhaMensal

    sigla = normalizaTexto(sigla)

    if planilhaMensal is not None:
        planilhaMensal.loc[planilhaMensal['ORGAO_ENTIDADE'].str.contains(palavra, case=False, na=False), 'SIGLA'] = sigla

def defineSiglas():
  #Gabinete e secretarias
  criaSigla('Gabinete do Prefeito', 'GBP')
  criaSigla('Gabinete do Vice-Prefeito', 'GVP')
  criaSigla('Casa Civil', 'CVL')
  criaSigla('Governo', 'SMG')
  criaSigla('Coordenação Governamental', 'SMCG')
  criaSigla('Fazenda', 'SMF')
  criaSigla('Integridade e Transparência', 'SMIT')
  criaSigla('Desenvolvimento Urbano', 'SMDU')
  criaSigla('Desenvolvimento Econômico', 'SMDE')
  criaSigla('Infraestrutura', 'SMI')
  criaSigla('Transportes', 'SMTR')
  criaSigla('Conservação', 'SECONSERVA')
  criaSigla('Educação', 'SME')
  criaSigla('Assistência Social', 'SMAS')
  criaSigla('Saúde', 'SMS')
  criaSigla('Administração', 'SMA')
  criaSigla('Trabalho e Renda', 'SMTE')
  criaSigla('Cultura', 'SMC')
  criaSigla('Pessoa com Deficiência', 'SMPD')
  criaSigla('Meio Ambiente e Clima', 'SMAC')
  criaSigla('Esportes', 'SMEL')
  criaSigla('Habitação', 'SMH')
  criaSigla('Ciência, Tecnologia e Inovação', 'SMCT')
  criaSigla('Envelhecimento Saudável e Qualidade de Vida', 'SEMESQV')
  criaSigla('Ordem Pública', 'SEOP')
  criaSigla('Proteção e Defesa dos Animais', 'SMPDA')
  criaSigla('Turismo', 'SMTUR-RIO')
  criaSigla('Proteção e Defesa do Consumidor', 'SEDECON')
  criaSigla('Políticas para Mulheres e Cuidados', 'SPM-RIO')
  criaSigla('Juventude Carioca', 'JUV-RIO')
  criaSigla('Ação Comunitária', 'SEAC-RIO')
  criaSigla('Cidadania e Família', 'SECID')
  criaSigla('Integração Metropolitana', 'SEIM')
  criaSigla('Economia Solidária', 'SES-RIO')
  criaSigla('Inclusão', 'SEI-RIO')
  criaSigla('Direitos Humanos e Igualdade Racial', 'SEDHIR')
  criaSigla('Controladoria Geral', 'CGM-Rio')
  criaSigla('Procuradoria Geral', 'PGM')

  #Institutos, fundações e empresas
  criaSigla('Previdência e Assistência', 'PREVI-RIO')
  criaSigla('Urbanismo Pereira Passos', 'IPP')
  criaSigla('Guarda Municipal', 'GM-RIO')
  criaSigla('Geotécnica', 'GEO-RIO')
  criaSigla('Águas do Município', 'RIO-ÁGUAS')
  criaSigla('Parques e Jardins', 'FPJ')
  criaSigla('Planetário', 'PLANETÁRIO')
  criaSigla('Jardim Zoológico', 'RIO-ZOO')
  criaSigla('Cidade das Artes', 'CIDADE DAS ARTES')
  criaSigla('Multimeios', 'MULTIRIO')
  criaSigla('Distribuidora de Filmes', 'RIOFILME')
  criaSigla('Informática', 'IPLANRIO')
  criaSigla('Artes Gráficas', 'IMPRENSA DA CIDADE')
  criaSigla('Parcerias e Investimentos', 'CCPAR')
  criaSigla('Urbanização', 'RIO-URBE')
  criaSigla('Turismo do Município', 'RIOTUR')
  criaSigla('Pública de Saúde', 'RIOSAÚDE')
  criaSigla('Energia e Iluminação', 'RIOLUZ')
  criaSigla('Limpeza Urbana', 'COMLURB')
  criaSigla('Engenharia de Tráfego', 'CET-RIO')
  criaSigla('Transportes Coletivos', 'CMTC Rio')
  criaSigla('Feiras, Exposições e Congressos', 'RIOCENTRO')
  criaSigla('Fomento do Município', 'INVEST.RIO')

def processaPlanilhas():
    global planilhaMensal, planilhaMinibio

    if planilhaMensal is None or planilhaMinibio is None:
        return None, None, None

    pd.set_option('display.max_rows', None)
    planilhaMensal['INICIO_LOTACAO'] = planilhaMensal['INICIO_LOTACAO'].dt.strftime('%d/%m/%Y')
    planilhaMinibio['ORGAO_ENTIDADE'] = planilhaMinibio['ORGAO_ENTIDADE'].apply(normalizaTexto)

    planilhaMensal = planilhaMensal[planilhaMensal['NOME'].isin(planilhaMinibio['NOME'])]
    nomesDuplicados = planilhaMensal[planilhaMensal.duplicated(subset=['NOME'], keep=False)]

    defineSiglas()

    planilhasMescladas = planilhaMensal.merge(planilhaMinibio, on = 'NOME', how = 'inner', suffixes = ('_MENSAL', '_MINIBIO'))
    planilhasMescladas['IGUAIS'] = planilhasMescladas['ORGAO_ENTIDADE_MINIBIO'] == planilhasMescladas['SIGLA']

    return nomesDuplicados, planilhasMescladas, planilhaMensal

def processaHistorico():
    global planilhaErgon, planilhaHistorico

    if planilhaErgon is None or planilhaHistorico is None:
        return None

    valoresDuplicados = planilhaErgon[planilhaErgon.duplicated(subset=['CPF'], keep=False)]
    return valoresDuplicados