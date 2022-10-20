# %%
import pandas as pd
import os
from datetime import date
import logging

# Pega a data do dia e transforma em uma string com formato em ano, mês e dia unificado ex: 20221011
DateTodayStr = '{}{}{}'.format(date.today().year, date.today().month, date.today().day)

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
                                          converters={'EMAIL': str.upper, 'NOME': str.upper}, sep=',',
                                          encoding='ISO-8859-1')
                log.debug(f'O DataFrame df_protheus recebeu o arquivo {i}')

            elif 'RM' in i:
                df_rm = pd.read_csv(i, converters={'Email': str.upper, 'Nome': str.upper}, sep=',',
                                    encoding='ISO-8859-1')
                log.debug(f'O DataFrame df_rm recebeu o arquivo {i}')
    except:
        log.error('Não foi possível localizar os arquivos do dia {}'.format(date.today()))

# Cria uma nova coluna nos DATA FRAMES com a empresa correspondente
try:
    df_protheus['EMPRESA'] = 'Rio Quente'
    df_rm['EMPRESA'] = 'Sauipe'
except Exception as error:
    log.error(error)

# Converte em maiusculo os titulos de todas as colunas
df_protheus.columns = df_protheus.columns.str.upper()
df_rm.columns = df_rm.columns.str.upper()

# %%
# Remove duplicadas mantendo ultimas ocorrências em cada arquivo
df_protheus.drop_duplicates(subset='CPF', keep='last', inplace=True)
df_rm.drop_duplicates(subset='CPF', keep='last', inplace=True)

# %%
# Remove as colunas indesejadas dos dfs, mantendo somente as necessárias
for i in df_protheus.columns:
    if i != 'MATRICULA' and i != 'EMPRESA' and i != 'FILIAL' and i != 'CARGO' and i != 'NOME' and i != 'CPF' and i != 'EMAIL' and i != 'SITUACAO' and i != 'DATA_DEMISSAO' and i != 'DATA_ADMISSAO':
        # print(i)
        df_protheus.drop(i, axis=1, inplace=True)
for i in df_rm.columns:
    if i != 'MATRICULA' and i != 'EMPRESA' and i != 'FILIAL' and i != 'CARGO' and i != 'NOME' and i != 'CPF' and i != 'EMAIL' and i != 'SITUACAO' and i != 'DATA_DEMISSAO' and i != 'DATA_ADMISSAO':
        # print(i)
        df_rm.drop(i, axis=1, inplace=True)

# %%
# Concatenar os dois dataFrames
df_auxiliar = pd.concat([df_protheus, df_rm])

# Altera os valores descritos como Demitidos para Demitido, para padronizar
df_auxiliar.loc[df_auxiliar['SITUACAO'] == 'Demitidos', 'SITUACAO'] = 'Demitido'

# CPF vázios deixam de ter valor nan para string vázia
df_auxiliar['CPF'] = df_auxiliar['CPF'].fillna('')

# Datas com valor nan passam a ter valor igual a 0
df_auxiliar['DATA_DEMISSAO'] = df_auxiliar['DATA_DEMISSAO'].fillna('')

# Datas de demissão que recebem a string '/  /' recebem 0 também para padronizar como vazias (igual linha acima)
df_auxiliar.loc[df_auxiliar['DATA_DEMISSAO'] == '/  /', 'DATA_DEMISSAO'] = ''

# Ordena as linhas por valores de CPFs
df_auxiliar.sort_values(by='CPF', inplace=True)

# Redefine o index para ficarm em ordem de 0 a n
# Desta forma as duplicatas entre empresas ficam em sequencia
df_auxiliar.reset_index(inplace=True, drop=True)
print(df_auxiliar.index)
# df_auxiliar.sort_index()

# %%
# Cria um novo df somente com as informações de colunas iguas o df concatenado
df_final = pd.DataFrame(columns=df_auxiliar.columns)
print(df_final)

# %%
# Especifica que a coluna DATA_DEMISSAO é do tipo date
df_auxiliar['DATA_DEMISSAO'] = pd.to_datetime(df_auxiliar['DATA_DEMISSAO'], errors='ignore', dayfirst=True)

# %% Remove duplicadas mantendo somente as linhas com SITUACAO diferente de Demitido ou, quando SITUACAO igual para
# ambas, manter apenas uma, dando preferência a ultima ocorrência considerando a data de demissao
i = 0
tamanho = len(df_auxiliar.index)
# while i < len(df_auxiliar.index) - 1:
try:
    while True:
        print(f'Começando linha {i}')
        if i < len(df_auxiliar.index) - 1:
            if df_auxiliar.loc[i, 'CPF'] == df_auxiliar.loc[i + 1, 'CPF']:
                print(f'\n{i} e {i + 1} são iguais')
                if df_auxiliar.loc[i, 'SITUACAO'] == df_auxiliar.loc[i + 1, 'SITUACAO']:
                    print(f'\nSITUACAO DE {i} e {i + 1} são iguais')
                    if df_auxiliar.loc[i, 'DATA_DEMISSAO'] != '' and df_auxiliar.loc[i + 1, 'DATA_DEMISSAO'] != '':
                        if pd.to_datetime(df_auxiliar.loc[i, 'DATA_DEMISSAO']) >= pd.to_datetime(
                                df_auxiliar.loc[i + 1, 'DATA_DEMISSAO']):
                            df_final = pd.concat([df_final, df_auxiliar.iloc[[i], 0:]])
                            print('A data de demissao de {} = {} é maior que {} = {}'.format(i, df_auxiliar.loc[
                                i, 'DATA_DEMISSAO'], i + 1, df_auxiliar.loc[i + 1, 'DATA_DEMISSAO']))
                            print('Inserindo {}'.format(df_auxiliar.loc[[i], :]))
                            i += 2
                        else:
                            df_final = pd.concat([df_final, df_auxiliar.iloc[[i + 1], 0:]])
                            print('A data de demissao de {} = {} é maior que {} = {}'.format(i + 1, df_auxiliar.loc[
                                i + 1, 'DATA_DEMISSAO'], i, df_auxiliar.loc[i, 'DATA_DEMISSAO']))
                            print('Inserindo {}'.format(df_auxiliar.loc[[i + 1], :]))
                            i += 2
                    else:
                        if df_auxiliar.loc[i, 'DATA_DEMISSAO'] == 0:
                            df_final = pd.concat([df_final, df_auxiliar.iloc[[i + 1], 0:]])
                            print('Inserindo {}'.format(df_auxiliar.loc[[i + 1], :]))
                            i += 2
                        else:
                            df_final = pd.concat([df_final, df_auxiliar.iloc[[i], 0:]])
                            print('Inserindo {}'.format(df_auxiliar.loc[[i], :]))
                            i += 2
                else:
                    print(f'\nSITUACAO DE {i} e {i + 1} NÃO são iguais')
                    if (df_auxiliar.loc[i, 'SITUACAO'] == 'Demitido') and (
                            df_auxiliar.loc[i + 1, 'SITUACAO'] != 'Demitido'):
                        df_final = pd.concat([df_final, df_auxiliar.iloc[[i + 1], 0:]])
                        print('Inserindo {}'.format(df_auxiliar.loc[[i + 1], :]))
                        i += 2
                        print(f'Almentando + 2 em i ficou {i}')
                    else:
                        df_final = pd.concat([df_final, df_auxiliar.iloc[[i], 0:]])
                        print('Inserindo {}'.format(df_auxiliar.loc[[i], :]))
                        i += 2
                        print(f'Almentando + 2 em i ficou {i}')
            else:
                df_final = pd.concat([df_final, df_auxiliar.iloc[[i], 0:]])
                print('Inser {}'.format(df_auxiliar.loc[[i], :]))
                i += 1
        elif i == len(df_auxiliar.index) - 1:
            df_final = pd.concat([df_final, df_auxiliar.iloc[[i], 0:]])
            print('Inserindo {}'.format(df_auxiliar.loc[[i], :]))
            break
except Exception as error:
    log.error(error)
else:
    log.debug('Eliminadas as ocorrências antigas de associados duplicados em ambos os sites Rio Quente e Sauipe')

# %%
# Salvar em csv
# Caso não exista uma pasta para armazenar os arquivos resultantes, ela será criada nesta execução do if
# Se a pasta já existir apenas salvamos o arquivo nesta pasta
if not os.path.exists('ProtheusRM'):
    os.makedirs('ProtheusRM')
    df_final.sort_values(by='CPF', inplace=True)
    df_final.reset_index(drop=True, inplace=True)
    #df_final['DATA_DEMISSAO'] = pd.to_datetime(df_final['DATA_DEMISSAO'], errors='ignore', dayfirst=True)
    try:
        df_final.to_csv('ProtheusRM/ProtheusRM_Final_' + DateTodayStr + '.csv', header=True, sep=';', encoding='UTF-8')
    except Exception as error:
        log.error(error)
    else:
        log.debug('O arquivo resultante ProtheusRM_Final_' + DateTodayStr + '.csv foi criado com êxito')
else:
    df_final.sort_values(by='CPF', inplace=True)
    df_final.reset_index(drop=True, inplace=True)
    #df_final['DATA_DEMISSAO'] = pd.to_datetime(df_final['DATA_DEMISSAO'], errors='ignore', dayfirst=True)
    try:
        df_final.to_csv('ProtheusRM/ProtheusRM_Final_' + DateTodayStr + '.csv', header=True, sep=';', encoding='UTF-8')
    except Exception as error:
        log.error(error)
    else:
        log.debug('O arquivo resultante ProtheusRM_Final_' + DateTodayStr + '.csv foi criado com êxito')
