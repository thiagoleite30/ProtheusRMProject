# %%
import pandas as pd
import numpy as np
import os
from datetime import date
import logging

# Pega a data do dia e transforma em uma string com formato em ano, mês e dia unificado ex: 20221011
DateTodayStr = '{}{:02}{:02}'.format(date.today().year, date.today().month, date.today().day)

# Cria uma pasta para armezenar logs de execução diários
# Definimos um formato de menssagem de log e criamos o log
if not os.path.exists('Logs'):
    os.makedirs('Logs')
LOG_FORMAT = '%(levelname)s %(asctime)s - %(message)s'
logging.basicConfig(filename='Logs/log_' + DateTodayStr + '.log', level=logging.DEBUG, filemode='w', format=LOG_FORMAT)
log = logging.getLogger()

# Listando arquivos do diretorio em uma lista com todos os arquivos
try:
    AllFilesPath = os.listdir()
except:
    log.error('Erro ao listar os arquivos do diretorio atual ' + os.getcwd())
else:
    log.debug('Listagem de arquivos do diretorio ' + os.getcwd() + ' concluída com êxito')

# Pega somente os arquivos do dia e colocam na lista vazia criada anteriormente
for i in AllFilesPath:
    try:
        if 'Protheus_' + DateTodayStr in i or 'RM_' + DateTodayStr in i and '.~' not in i:
            # Identifica qual arquivo é do Protheus e qual é do RM e já transforma-os em dataframes
            # Converte todos os calores das colunas EMAIL e NOME em maiúsculo
            # Altera a forma de separação dos arquivos para , e a codificação para 'ISO-8859-1'
            if 'Protheus' in i:
                df_protheus = pd.read_csv(i,
                                          converters={'EMAIL': str.upper, 'NOME': str.upper}, sep=';',
                                          encoding='ISO-8859-1')
                log.debug(f'O DataFrame df_protheus recebeu o arquivo {i}')

            elif 'RM' in i:
                df_rm = pd.read_csv(i,
                                          converters={'EMAIL': str.upper, 'NOME': str.upper}, sep=';',
                                          encoding='ISO-8859-1')
                log.debug(f'O DataFrame df_rm recebeu o arquivo {i}')
    except:
        log.error('Não foi possível localizar os arquivos do dia {}'.format(date.today()))

# Trabalha ambos os DataFrame para obter um geral rotulado
try:
    #Removendo duplicidades de CPF em cada um dos DataFrames mantendo a ultima ocorrência
    #- Desta forma, em caso de uma pessoa ter sido contratada mais de uma vez em cargos diferentes, mantem-se apenas a ultima ocorrência.

    df_protheus.drop_duplicates(subset='CPF', keep='last', inplace=True)
    df_rm.drop_duplicates(subset='CPF', keep='last', inplace=True)

    # Converte em maiusculo os titulos de todas as colunas
    df_protheus.columns = df_protheus.columns.str.upper()
    df_rm.columns = df_rm.columns.str.upper()

    #Removendo algumas colunas que não precisaremos utilizar
    df_protheus.drop(['EMPRESA_CNPJ', 'SETOR', 'COD_CARGO', 'VINCULO', 'TELEFONE', 'CELULAR', 'EMAIL'], axis=1, inplace=True)
    df_rm.drop(['EMPRESA_CNPJ', 'SETOR', 'VINCULO', 'TELEFONE', 'CELULAR', 'EMAIL'], axis=1, inplace=True)

    #Criando coluna para diferenciar a Empresa
    df_protheus['EMPRESA'] = 'Rio Quente'
    df_rm['EMPRESA'] = 'Sauipe'

    df_rm.rename(columns={'SITUACAO_DATA_INICIO' : 'SITUACAO_DATA_INICIO_AFAST', 'SITUACAO_DATA_FIM' : 'SITUACAO_DATA_FIM_AFAST'}, inplace=True)

    #  No RM o valor de Situação 'Demitido' estava escrito como Demitidos, então igualamos ao Protheus como 'Demitido'
    df_rm.loc[df_rm['SITUACAO'] == 'Demitidos', 'SITUACAO'] = 'Demitido'

    # No DataFrame do Protheus haviam valores de datas apenas com '/  /' e '//' então convertemos em valores NaN
    df_protheus['DATA_DEMISSAO'].replace(['/  /', '//'], np.nan, inplace=True)
    df_protheus['DATA_ADMISSAO'].replace(['/  /', '//'], np.nan, inplace=True)
    df_protheus['SITUACAO_DATA_INICIO_AFAST'].replace(['/  /', '//'], np.nan, inplace=True)
    df_protheus['SITUACAO_DATA_FIM_AFAST'].replace(['/  /', '//'], np.nan, inplace=True)

except Exception as error:
    log.error(error)

# Lendo o arquivo de cargos rotulados e inserindo informações de matricula e cargo gestor:
try:
    df_protheus_rotulados = pd.read_excel('df_protheus_rotulados.xlsx', index_col=0)

    # Inserindo colunas com informações de matricula e cargo no gestor no arquivo rotulado
    for idx, row in df_protheus_rotulados.iterrows():
        if row['GESTOR'] in df_protheus['NOME'].values:
            # print(idx,row['GESTOR'])
            df_protheus_rotulados.loc[idx, 'MATRICULA_GESTOR'] = \
                df_protheus[df_protheus['NOME'] == row['GESTOR']]['MATRICULA'].values[0]
            df_protheus_rotulados.loc[idx, 'CARGO_GESTOR'] = \
                df_protheus[df_protheus['NOME'] == row['GESTOR']]['CARGO'].values[0]

    # Salvando arquivo

    df_protheus_rotulados.to_excel('df_protheus_rotulados.xlsx', header=True)
except Exception as error:
    log.error(error)
else:
    log.debug('Arquivo de cargos rotulados lido e novas informações de matricula e cargo do gesto inseridas. Arquivo salvo como df_protheus_rotulados.xlsx!')

print(df_protheus_rotulados)
# ROTULANDO O ARQUIVO PROTHEUS
try:
    df_protheus.reset_index(drop=True, inplace=True)

    #Cria uma coluna temporária para comparar cargos com maior precisão. Evita problemas com cargo de mesmo nome e CCs diferentes
    df_protheus['CC_CARGO'] = df_protheus['CENTRO_CUSTO'] + '-' + df_protheus['CARGO']
    df_protheus_rotulados['CC_CARGO'] = df_protheus_rotulados['CENTRO_CUSTO'] + '-' + df_protheus_rotulados['CARGO']

    #FOR ITERROWS: AQUI SÃO CLASSIFICADOS OS CARGOS NO PROTHEUS RESULTANTE
    print('Iniciando o iterrows')
    print(f'Tamanho do df_protheus {len(df_protheus)}')
    try:
        df_protheus['GESTOR'] = np.nan
        for idx, row in df_protheus.iterrows():
            if row['CC_CARGO'] in df_protheus_rotulados['CC_CARGO'].to_list():
                # df_protheus_rotulados['CC_CARGO'].to_list().index(row['CC_CARGO'])
                #print(df_protheus.loc[idx, 'NOME'])
                #print(idx, df_protheus_rotulados.loc[df_protheus_rotulados['CC_CARGO'].to_list().index(row['CC_CARGO']), 'GESTOR'])
                indice = df_protheus_rotulados['CC_CARGO'].to_list().index(row['CC_CARGO'])
                df_protheus.loc[idx, 'GESTOR'] = df_protheus_rotulados.loc[indice, 'GESTOR']
            #else:
                #print(idx, np.nan)

        #INSERINDO INFORMAÇÕES DE MATRICULA E CARGO DO GESTOR NO DATAFRAME DO PROTHEUS
        for idx, row in df_protheus.iterrows():
            if row['GESTOR'] in df_protheus['NOME'].values:
                # print(idx,row['GESTOR'])
                df_protheus.loc[idx, 'MATRICULA_GESTOR'] = \
                df_protheus[df_protheus['NOME'] == row['GESTOR']]['MATRICULA'].values[0]
                df_protheus.loc[idx, 'CARGO_GESTOR'] = df_protheus[df_protheus['NOME'] == row['GESTOR']]['CARGO'].values[0]
    except Exception as error:
        log.error(error)
    else:
        print(f'Tamanho do df_protheus {len(df_protheus)}')

except Exception as error:
    log.error(error)
else:
    log.debug('DataFrame df_protheus rotulado com sucesso!')

# %%
# Concatenar os dois dataFrames
try:
    print(df_protheus.columns)
    df_ProtheusRM = pd.concat([df_protheus, df_rm], axis=0)

    #Mudando ordem das colunas para ficar melhor de visualizar
    ordemCols = ['EMPRESA', 'FILIAL', 'MATRICULA', 'NOME', 'CPF', 'DATA_ADMISSAO', 'CENTRO_CUSTO', 'CARGO',
                 'SITUACAO', 'SITUACAO_DATA_INICIO_AFAST', 'SITUACAO_DATA_FIM_AFAST',
                 'DATA_DEMISSAO', 'GESTOR', 'MATRICULA_GESTOR', 'CARGO_GESTOR']
    df_ProtheusRM = df_ProtheusRM[ordemCols]

    #Mudando o tipo de colunas para DateTime para as colunas que carregam informações de Datas
    df_ProtheusRM['SITUACAO_DATA_INICIO_AFAST'] = pd.to_datetime(df_ProtheusRM['SITUACAO_DATA_INICIO_AFAST'],
                                                                dayfirst=True).apply(lambda x: x.replace(tzinfo=None))
    df_ProtheusRM['SITUACAO_DATA_FIM_AFAST'] = pd.to_datetime(df_ProtheusRM['SITUACAO_DATA_FIM_AFAST'],
                                                             dayfirst=True).apply(lambda x: x.replace(tzinfo=None))
    df_ProtheusRM['DATA_ADMISSAO'] = pd.to_datetime(df_ProtheusRM['DATA_ADMISSAO'], dayfirst=True).apply(
        lambda x: x.replace(tzinfo=None))
    df_ProtheusRM['DATA_DEMISSAO'] = pd.to_datetime(df_ProtheusRM['DATA_DEMISSAO'], dayfirst=True).apply(
        lambda x: x.replace(tzinfo=None))

    # Ordenando por CPF para que os duplicados entre sites fiquem um sobre o outro
    df_ProtheusRM.sort_values(by='CPF', inplace=True)

    # Resetando o index do DataFrame maior para trabalhar no for iterrow
    df_ProtheusRM.reset_index(drop=True, inplace=True)

    ### Removendo CPFs não informados
    df_ProtheusRM.dropna(subset=['CPF'], inplace=True)

    # Obtendo um filtro com os duplicados
    # Precisa manter o keep='last' para que a primeira ocorrência de duplicadas
    filtro = df_ProtheusRM.duplicated(subset='CPF', keep='last')

    for idx, row in filtro.iteritems():
        if row:
            print(idx, row)
except Exception as error:
    log.error(error)
else:
    log.debug('DataFrames df_protheus e df_rm concatenados com sucesso!')

#Após obtenção do filtro
#- Primeiro verificamos as ocorrências de datas vázias (com valor NaT) e as mantemos, pois assumem que são associados 'Ativos' (podendo ser 'Férias', 'Afastados' e etc).
#- Em seguida onde o associada constam com data de demissão existente para ambas as duplicadas, optamos por manter o último periodo observado.
try:
    for idx, row in filtro.iteritems():
        if row:
            if isinstance(df_ProtheusRM.loc[idx, 'DATA_DEMISSAO'], pd._libs.tslibs.nattype.NaTType):
                df_ProtheusRM.drop(idx + 1, inplace=True)
            elif isinstance(df_ProtheusRM.loc[idx + 1, 'DATA_DEMISSAO'], pd._libs.tslibs.nattype.NaTType):
                df_ProtheusRM.drop(idx, inplace=True)
            elif df_ProtheusRM.loc[idx, 'DATA_DEMISSAO'] >= df_ProtheusRM.loc[idx + 1, 'DATA_DEMISSAO']:
                df_ProtheusRM.drop(idx + 1, inplace=True)
            elif df_ProtheusRM.loc[idx, 'DATA_DEMISSAO'] < df_ProtheusRM.loc[idx + 1, 'DATA_DEMISSAO']:
                df_ProtheusRM.drop(idx, inplace=True)
except Exception as error:
    log.error(error)
else:
    log.debug('Eliminadas as ocorrências antigas de associados duplicados em ambos os sites Rio Quente e Sauipe. Mantidos os Ativos (férias, afastamento e etc)!')

# %%
# Salvar em csv
# Caso não exista uma pasta para armazenar os arquivos resultantes, ela será criada nesta execução do if
# Se a pasta já existir apenas salvamos o arquivo nesta pasta
if not os.path.exists('ProtheusRM'):
    os.makedirs('ProtheusRM')
    df_ProtheusRM.sort_values(by='CPF', inplace=True)
    df_ProtheusRM.reset_index(drop=True, inplace=True)
    #df_ProtheusRM['DATA_DEMISSAO'] = pd.to_datetime(df_ProtheusRM['DATA_DEMISSAO'], errors='ignore', dayfirst=True)
    try:
        df_ProtheusRM.to_csv('ProtheusRM/ProtheusRM_Final_' + DateTodayStr + '.csv', header=True, sep=';', encoding='UTF-8')
    except Exception as error:
        log.error(error)
    else:
        log.debug('O arquivo resultante ProtheusRM_Final_' + DateTodayStr + '.csv foi criado com êxito')
else:
    df_ProtheusRM.sort_values(by='CPF', inplace=True)
    df_ProtheusRM.reset_index(drop=True, inplace=True)
    #df_ProtheusRM['DATA_DEMISSAO'] = pd.to_datetime(df_ProtheusRM['DATA_DEMISSAO'], errors='ignore', dayfirst=True)
    try:
        df_ProtheusRM.to_csv('ProtheusRM/ProtheusRM_Final_' + DateTodayStr + '.csv', header=True, sep=';', encoding='UTF-8')
    except Exception as error:
        log.error(error)
    else:
        log.debug('O arquivo resultante ProtheusRM_Final_' + DateTodayStr + '.csv foi criado com êxito')
