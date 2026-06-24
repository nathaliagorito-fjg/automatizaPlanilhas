import pandas as pd
import unicodedata as uni
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    handlers=[
        logging.FileHandler("execucao_RPA.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

planilhaMensal = None
planilhaMinibio = None
planilhaErgon = None
planilhaHistorico = None
diretorioHistorico = None

def normalizaTexto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).upper().strip()
    texto = ''.join(c for c in uni.normalize('NFD', texto) if uni.category(c) != 'Mn')
    texto = texto.replace('-', '').replace('.', '')
    return texto

def padronizaColunas(df):

    mapeamento = {}
    for col in df.columns:
        col_upper = str(col).upper().strip()
        
        if col_upper in ['SETOR', 'NOME_SETOR', 'NOMESETOR']:
            mapeamento[col] = 'NOME_SETOR'
            
        elif col_upper in ['REFERENCIA', 'REF', 'SIMBOLO DO CARGO', 'SÍMBOLO DO CARGO', 'REFERÊNCIA']:
            mapeamento[col] = 'REFERENCIA'
            
        elif col_upper in ['NOME', 'NOME COMPLETO']:
            mapeamento[col] = 'NOME'
            
        elif col_upper in ['ORGAO', 'ORGAO_ENTIDADE', 'ORGAO/ENTIDADE', 'ÓRGÃO', 'SIGLA_ORGAO_ENTIDADE', 'ORGÃO_ENTIDADE']:
            mapeamento[col] = 'ORGAO_ENTIDADE'
            
    return df.rename(columns=mapeamento)

def criaSigla(palavra, sigla):
    global planilhaMensal
    sigla = normalizaTexto(sigla)
    if planilhaMensal is not None and 'ORGAO_ENTIDADE' in planilhaMensal.columns:
        planilhaMensal.loc[planilhaMensal['ORGAO_ENTIDADE'].str.contains(palavra, case=False, na=False), 'SIGLA'] = sigla

def defineSiglas():

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

def processaPlanilhasLideresMensais():
    global planilhaMensal, planilhaMinibio

    if planilhaMensal is None or planilhaMinibio is None:
        logging.warning("Processamento abortado: Planilhas de líderes mensais não foram carregadas.")
        return None, None, None

    try:
        logging.info("Iniciando o processamento de Líderes Mensais.")
        pd.set_option('display.max_rows', None)
        
        planilhaMensal = padronizaColunas(planilhaMensal)
        planilhaMinibio = padronizaColunas(planilhaMinibio)

        if 'INICIO_LOTACAO' in planilhaMensal.columns:
            planilhaMensal['INICIO_LOTACAO'] = pd.to_datetime(planilhaMensal['INICIO_LOTACAO'], errors='coerce')
            planilhaMensal['INICIO_LOTACAO'] = planilhaMensal['INICIO_LOTACAO'].dt.strftime('%d/%m/%Y')
        
        if 'ORGAO_ENTIDADE' in planilhaMinibio.columns:
            planilhaMinibio['ORGAO_ENTIDADE'] = planilhaMinibio['ORGAO_ENTIDADE'].apply(normalizaTexto)

        planilhaMensal = planilhaMensal[planilhaMensal['NOME'].isin(planilhaMinibio['NOME'])]
        nomesDuplicados = planilhaMensal[planilhaMensal.duplicated(subset=['NOME'], keep=False)]

        planilhaMensal['SIGLA'] = ""
        defineSiglas()
        planilhaMensal['SIGLA'] = planilhaMensal['SIGLA'].fillna("")

        planilhasMescladas = planilhaMensal.merge(
            planilhaMinibio, 
            on='NOME', 
            how='inner', 
            suffixes=('_MENSAL', '_MINIBIO')
        )

        colunas_validacao = [
            'ORGAO_ENTIDADE_MINIBIO', 'SIGLA', 
            'REFERENCIA_MENSAL', 'REFERENCIA_MINIBIO', 
            'NOME_SETOR_MENSAL', 'NOME_SETOR_MINIBIO'
        ]
        for col in colunas_validacao:
            if col in planilhasMescladas.columns:
                planilhasMescladas[col] = planilhasMescladas[col].fillna("").astype(str).str.strip()
            else:
                planilhasMescladas[col] = ""

        diverge_orgao = planilhasMescladas['ORGAO_ENTIDADE_MINIBIO'] != planilhasMescladas['SIGLA']
        diverge_referencia = planilhasMescladas['REFERENCIA_MENSAL'] != planilhasMescladas['REFERENCIA_MINIBIO']
        diverge_setor = planilhasMescladas['NOME_SETOR_MENSAL'] != planilhasMescladas['NOME_SETOR_MINIBIO']

        planilhasDivergentes = planilhasMescladas[diverge_orgao | diverge_referencia | diverge_setor]

        colunas_foco = [
            'NOME', 'ORGAO_ENTIDADE_MENSAL', 'ORGAO_ENTIDADE_MINIBIO', 'SIGLA',
            'REFERENCIA_MENSAL', 'REFERENCIA_MINIBIO',
            'NOME_SETOR_MENSAL', 'NOME_SETOR_MINIBIO'
        ]
        
        colunas_existentes = [col for col in colunas_foco if col in planilhasDivergentes.columns]
        

        planilhasDivergentes = planilhasDivergentes.drop_duplicates(subset=colunas_existentes)

        logging.info(f"Processamento concluído. Encontrados {len(planilhasDivergentes)} registros divergentes.")
        return nomesDuplicados, planilhasDivergentes, planilhaMensal

    except Exception as e:
        logging.error(f"Falha crítica no processamento de Líderes Mensais: {str(e)}", exc_info=True)
        raise e

def processaPlanilhasHistoricoMinibio():
    global planilhaErgon, planilhaHistorico

    if planilhaErgon is None or planilhaHistorico is None:
        logging.warning("Processamento abortado: Planilhas do Histórico não foram carregadas.")
        return None

    try:
        logging.info("Iniciando o processamento do Histórico da Minibio.")
        valoresDuplicados = planilhaErgon[planilhaErgon.duplicated(subset=['CPF'], keep=False)]
        return valoresDuplicados
    except Exception as e:
        logging.error(f"Falha crítica no processamento do Histórico: {str(e)}", exc_info=True)
        raise e