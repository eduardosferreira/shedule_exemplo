#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
MODULO ...: TESHUVA
SCRIPT ...: retorna_dados_banco.py
CRIACAO ..: 11/02/2022
AUTOR ....: Victor Santos Cardoso - KYROS TECNOLOGIA
DESCRICAO.: 
----------------------------------------------------------------------------------------------
  HISTORICO : 
----------------------------------------------------------------------------------------------
"""
from asyncio import constants
from concurrent.futures import thread
import json
import os
import sys
import threading
import time
import datetime

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)

import configuracoes
import comum
import layout
import sql
import conversor.classe_converte_conv115_to_json
import conversor.converte_conv86_to_json 
import fnmatch
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

log.gerar_log_em_arquivo = True
comum.carregaConfiguracoes(configuracoes)
layout.carregaLayout()

status_threads                      = []
retorno_threads                     = []
lista_arquivos_conv115              = []
lista_notas                         = {}

def processar():
    
    global status_threads,retorno_threads,lista_arquivos_conv115,lista_notas

    diretorio_arquivos_conv86 = configuracoes.diretorio_conv86
    diretorio_arquivos_conv86 = diretorio_arquivos_conv86.replace( '<<UF>>', configuracoes.uf )
    diretorio_arquivos_conv86 = diretorio_arquivos_conv86.replace( '<<DIR_PLEITO>>', configuracoes.dir_pleito )

    for arquivo in os.listdir(diretorio_arquivos_conv86):
        if fnmatch.fnmatch(arquivo,'*.txt'):
            path_arquivo_conv86 = arquivo
    
    data_atua            = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    dados_arquivo_conv86 = conversor.converte_conv86_to_json.converte_conv86_to_json(os.path.join(diretorio_arquivos_conv86,path_arquivo_conv86 ))

    if not dados_arquivo_conv86:
        return False
    
    log('INICIANDO CRIAÇÃO DE LISTA DE NOTAS PARA PROCESSAMENTO')
    for registro in dados_arquivo_conv86[0]['item']:

        data_emissao = registro['DATA_EMISSAO'][:4] + registro['DATA_EMISSAO'][4:6]

        chave_nota = data_emissao, registro['SERIE'], registro['MODELO']

        lista = []

        if chave_nota in lista_notas:
    
            lista = lista_notas[chave_nota]

        lista.append([registro['NUMERO_NF'], registro['NUM_ITEM']])

        lista_notas[chave_nota] = lista
    
    for x in lista_notas.keys():

        lista_notas[x].sort()

    log('FINALIZANDO CRIAÇÃO DE LISTA DE NOTAS')

    dados_tipo_2_conv86 = dados_arquivo_conv86   
    
    lista_conv_115                = []
    resultado_tipo_2_conv86       = []
    resultado_divergencias        = []
    resultado_sem_correspondentes = []
    
    threads                             = [0 for i in range(int(configuracoes.maximo_threads))]
    status_threads                      = [0 for i in range(int(configuracoes.maximo_threads))]
    retorno_threads                     = [0 for i in range(int(configuracoes.maximo_threads))]

    lista_chaves_notas = [x for x in lista_notas.keys()]    

    while lista_chaves_notas:
    
        time.sleep(3)

        for X in range(0, len(status_threads)):
            
            status = status_threads[X] 

            if status == 0:
                
                if lista_chaves_notas:
                    chave = lista_chaves_notas.pop(0)

                    th = threading.Thread( target = get_all_conv115, args = ( X, {chave:lista_notas[chave]}))
                    th.start()
                    threads[X] = th
                    time.sleep(2)  

            if status == 2: 
                
                time.sleep(1)

                lista_conv_115 += retorno_threads[X]

                log('THREAD %s FINALIZADA'%X)
                status_threads[X] = 0 
            
            if status == 3:
                log(' ERRO '.center(100,'=')) 
                return False  

    esperar = True

    log('AGUARDANDO TODAS AS THREAD FINALIZAREM')

    while esperar:

        time.sleep(1)

        esperar = False
        
        for X in range(0, int(configuracoes.maximo_threads)):

            status = status_threads[X]

            if status > 0:
                esperar = True

            if status == 1:
                if not threads[X].is_alive():
                    
                    time.sleep(1)
                    
                    if status_threads[X] == 1:

                        status_threads[X] = 3

            if status == 2:

                lista_conv_115 += retorno_threads[X]
                
                log('THREAD %s FINALIZADA'%X) 
                
                status_threads[X] = 0

            if status == 3:
                log(' ERRO '.center(100,'=')) 
                return False 


    log('TODAS AS THREAD FINALIZARAM')

    retorno_result,retorno_divergencias,retorno_sem_correspondentes = cruza_dados(dados_tipo_2_conv86, lista_conv_115) 

    resultado_tipo_2_conv86       += retorno_result
    resultado_divergencias        += retorno_divergencias
    resultado_sem_correspondentes += retorno_sem_correspondentes
    
    log('ORDENANDO DADOS DA LISTA FINAL DO CONV86.')            
    resultado = comum.ordenaListaDicionarios( resultado_tipo_2_conv86, ['numero_linha'] )
    log('FIM DA ORDENAÇÃO')

    v_dir_base  = os.path.join(configuracoes.dir_geracao_arquivos, configuracoes.uf, configuracoes.dir_pleito)    

    log('INICIANDO A GERAÇÃO DOS ARQUIVOS CONV86, DIVERGENTES E SEM CORRESPONDENTES')
    if not os.path.isdir(v_dir_base) :
        log('CRIANDO DIRETÓRIO'.center(100,'='))
        log(v_dir_base)
        os.makedirs(v_dir_base)

    log(' NOME DO ARQUIVO', path_arquivo_conv86)
    arquivo_novo      = path_arquivo_conv86.split('.')[0] + '_' + data_atua +'.txt'
    v_nome_relatorio  = os.path.join(v_dir_base, arquivo_novo)
    arquivo_teshuva   = open(v_nome_relatorio, 'w', encoding=tipoArquivo(os.path.join(diretorio_arquivos_conv86,path_arquivo_conv86)))

    log(' ARQUIVO TESHUVA - CONV86:    ', v_nome_relatorio)

    primeira_linha = layout.geraLinha( dados_arquivo_conv86[0], 'LayoutConv86_1')

    arquivo_teshuva.write(primeira_linha + '\n')

    log('INICIANDO CARGA DO ARQUIVO FINAL')

    for registro in resultado:

        linha = {}
        for field in ['TIPO', 'MODELO', 'NUMERO_NF', 'SERIE', 'DATA_EMISSAO', 'HASH_CODE_ARQ', 'CNPJ_CPF', 'IE', 'RAZAOSOCIAL', 'CADG_COD', 'VALOR_TOTAL',\
                      'BASE_ICMS', 'VALOR_ICMS', 'NUM_ITEM', 'VALOR_ITEM', 'ICMS_ESTORNO', 'HIPOTESE_ESTORNO', 'MOTIVO_ESTORNO', 'NUM_RECLAMACAO']:
            linha[field] = registro[field]

        linha = layout.geraLinha(linha, 'LayoutConv86_2')
        arquivo_teshuva.write(linha + '\n')

    arquivo_teshuva.close()

    log('CARGA ARQUIVO FINAL CONCLUÍDA')

    arquivo_novo_div     = path_arquivo_conv86.split('.')[0] + '_DIVERGENCIAS_'+ data_atua +'.csv'
    v_nome_relatorio_div = os.path.join(v_dir_base, arquivo_novo_div)
    arquivo_divergencia  = open(v_nome_relatorio_div, 'w', encoding=tipoArquivo(os.path.join(diretorio_arquivos_conv86,path_arquivo_conv86)))

    log(' ARQUIVO DIVERGENCIA:         ', v_nome_relatorio_div)

    log('INICIANDO CARGA DO ARQUIVO DIVERGENCIA')
    
    lista_field = [ 'numero_linha', 'TIPO', 'MODELO', 'NUMERO_NF', 'NUMERO_NF_115', 'SERIE', 'SERIE_115', 'DATA_EMISSAO', 'DATA_EMISSAO_115', 'HASH_CODE_ARQ', 'HASH_COD_NF_115', 'CNPJ_CPF', 'CNPJ_CPF_115', 
    'IE', 'IE_115', 'RAZAOSOCIAL', 'RAZAO_SOCIAL_115', 'CADG_COD', 'CADG_COD_115', 'VALOR_TOTAL', 'VALOR_TOTAL_115', 'BASE_ICMS', 'BASE_ICMS_115', 'VALOR_ICMS',
    'VALOR_ICMS_115', 'NUM_ITEM', 'NUM_ITEM_115_ITEM', 'VALOR_ITEM','VALOR_TOTAL_115_ITEM', 'ICMS_ESTORNO', 'HIPOTESE_ESTORNO', 'MOTIVO_ESTORNO', 'NUM_RECLAMACAO' ]
    
    arquivo_divergencia.write(';'.join(x for x in lista_field ) + '\n' ) 
    
    resultado_div = comum.ordenaListaDicionarios( resultado_divergencias, ['numero_linha'] )

    for registro in resultado_div:
        
        linha = []
        
        for field in lista_field:
            try:
                if field.__contains__('VALOR') or field.__contains__('BASE'):
                    registro[field] = (float(registro[field]) / 100)
                
                linha.append(str(registro[field]))
            
            except Exception as e:
                print(registro)
                raise e

        arquivo_divergencia.write(';'.join(x for x in linha ) + '\n' )

    arquivo_divergencia.close()

    log('CARGA ARQUIVO DIVERGENCIA CONCLUÍDA')    

    arquivo_novo_corr          = path_arquivo_conv86.split('.')[0] + '_SEM_CORRESPONDENTE_'+ data_atua +'.csv'
    v_nome_relatorio_corr      = os.path.join(v_dir_base, arquivo_novo_corr)
    arquivo_sem_correspondente = open(v_nome_relatorio_corr, 'w', encoding=tipoArquivo(os.path.join(diretorio_arquivos_conv86,path_arquivo_conv86))) 
        
    log(' ARQUIVO SEM CORRESPONDENTE: ', v_nome_relatorio_corr)

    log('INICIANDO CARGA DO ARQUIVO SEM CORRESPONDENTE')

    lista_field_corr = ['numero_linha', 'TIPO', 'MODELO', 'NUMERO_NF', 'SERIE', 'DATA_EMISSAO', 'HASH_CODE_ARQ', 'CNPJ_CPF', 'IE', 'RAZAOSOCIAL', 'CADG_COD', 'VALOR_TOTAL', 
                        'BASE_ICMS', 'VALOR_ICMS', 'NUM_ITEM', 'VALOR_ITEM', 'ICMS_ESTORNO', 'HIPOTESE_ESTORNO', 'MOTIVO_ESTORNO', 'NUM_RECLAMACAO' ]     

    arquivo_sem_correspondente.write(';'.join(x for x in lista_field_corr ) + '\n' )

    resultado_corr = comum.ordenaListaDicionarios( resultado_sem_correspondentes, ['numero_linha'] )

    for registro in resultado_corr:
        linha = []

        for field in lista_field_corr:
            
            if field.__contains__('VALOR') or field.__contains__('BASE'):
                    registro[field] = (float(registro[field]) / 100)

            linha.append(str(registro[field]))

        arquivo_sem_correspondente.write(';'.join(x for x in linha ) + '\n' )

    arquivo_sem_correspondente.close()

    log('CARGA ARQUIVO SEM CORRESPONDENTE CONCLUÍDA') 

    return True

def tipoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r', encoding='utf-8')
        fd.readline()
        fd.close()
    except :
        return 'iso-8859-1'
    return 'utf-8'
    
def get_all_conv115( X, lista_notas_86):

    global status_threads, retorno_threads

    log('[THD %s] INICIANDO THREAD  '%X)
    
    arquivo_aberto              = ''
    lista_retorno               = []  
    status_threads[X]           = 1
    arquivo_conv115_retornado   = []   

    try:
        for lista in lista_notas_86:
        
            uf      = configuracoes.uf
            ano     = lista[0][:4]
            mes     = lista[0][4:6]
            serie   = lista[1]

            lista_arquivos_conv115 = geraListaArquivosConv115( uf, ano, mes, serie ) 

            for lista_arquivos in lista_arquivos_conv115:
                
                for nota,num_item in lista_notas_86[lista]:

                    if int(nota) >= int(lista_arquivos[2]) and int(nota) <= int(lista_arquivos[3]):

                        diretorio_arquivos        =  lista_arquivos[7]
                        ano_referencia            =  lista_arquivos[0].strftime("%Y")
                        volume                    =  lista_arquivos[4]
                        arquivo_conv115_retornado =  lista_arquivos[5]
                        ini_volume                =  lista_arquivos[2]            
                        fim_volume                =  lista_arquivos[3]

                        if arquivo_conv115_retornado != arquivo_aberto:

                            log('ARQUIVO RETORNADO' , arquivo_conv115_retornado)
                            log('DIRETORIO        ' , diretorio_arquivos)
                            log('ANO REFERENCIA   ' , ano_referencia)
                            log('VOLUME           ' , volume)
                            log('INICIO VOLUME    ' , ini_volume)
                            log('FIM VOLUME       ' , fim_volume)

                            conv           = conversor.classe_converte_conv115_to_json.CONVERSOR()
                            dados_conv115  = conv.converte_conv115_to_json( diretorio_arquivos, ano_referencia, volume, lista_notas_86[lista], ini_volume, fim_volume )
                            arquivo_aberto = arquivo_conv115_retornado

                            if dados_conv115:
                                lista_retorno.append(dados_conv115[0])

    except Exception as e:
        
        status_threads[X]  = 3
        log('ERRO THREAD', X, e)
        return False 

    retorno_threads[X] = lista_retorno  
    status_threads[X]  = 2    

    return True 

def cruza_dados( lista_aprocessar, dados_conv115 ) :

    log('INICIANDO O CRUZAMENTO DE DADOS...')

    global status_threads,retorno_threads,lista_arquivos_conv115,lista_notas

    lista_divergentes         = []
    lista_sem_correspondentes = []
    lista_resultado           = []
    idx_controle              = 0
    idx_mestre                = 0
    idx_item                  = 0
    chave_reg_c115            = {}
    chave_reg_86              = {}
    idx                       = 0
    idx_86                    = 0 

    log('INICIANDO CRIAÇÃO DE CHAVES CONV115')
    for idx_controle in range(len(dados_conv115)):
        
        for idx_mestre in range(len(dados_conv115[idx_controle]['mestre'])):
        
            for idx_item in range(len(dados_conv115[idx_controle]['mestre'][idx_mestre]['item'])): 

                chave_nota = [dados_conv115[idx_controle]['mestre'][idx_mestre]['NUMERO_NF']                   +'|'+\
                              dados_conv115[idx_controle]['mestre'][idx_mestre]['SERIE']                       +'|'+\
                              dados_conv115[idx_controle]['mestre'][idx_mestre]['DATA_EMISSAO']                +'|'+\
                              dados_conv115[idx_controle]['mestre'][idx_mestre]['item'][idx_item]['NUMERO_NF'] +'|'+\
                              dados_conv115[idx_controle]['mestre'][idx_mestre]['item'][idx_item]['NUM_ITEM']]

                posicao_reg = [idx_controle,idx_mestre,idx_item]
            
                chave_reg_c115[idx] = chave_nota
                chave_reg_c115[idx].append(posicao_reg)  
                idx += 1

    chave_reg_c115_ordenada = []

    for i in sorted(chave_reg_c115, key = chave_reg_c115.get):
        chave_reg_c115_ordenada.append(chave_reg_c115[i])

    log('FINALIZANDO CRIAÇÃO DE CHAVES CONV115')

    # with open('conv115.txt', 'w') as temp_file:
    #     for item in chave_reg_c115_ordenada:
    #         temp_file.write("%s\n" % item)

    log('INICIANDO CRIAÇÃO DE CHAVES CONV86')
    lista_aprocessar = comum.ordenaListaDicionarios( lista_aprocessar[0]['item'], ['NUMERO_NF','DATA_EMISSAO', 'SERIE', 'NUM_ITEM'] )

    for idx_c86 in range(len(lista_aprocessar)):        
        
        chave_nota_86 = [lista_aprocessar[idx_c86]['NUMERO_NF']    +'|'+\
                         lista_aprocessar[idx_c86]['SERIE']        +'|'+\
                         lista_aprocessar[idx_c86]['DATA_EMISSAO'] +'|'+\
                         lista_aprocessar[idx_c86]['NUMERO_NF']    +'|'+\
                         lista_aprocessar[idx_c86]['NUM_ITEM']]

        posicao_reg_86 = idx_c86

        chave_reg_86[idx_86] = chave_nota_86
        chave_reg_86[idx_86].append(posicao_reg_86)
        idx_86 += 1


    chave_reg_86_ordenada = []
    for i in sorted(chave_reg_86, key = chave_reg_86.get):
        chave_reg_86_ordenada.append(chave_reg_86[i])

    # with open('conv86.txt', 'w') as temp_file:
    #     for item in chave_reg_86_ordenada:
    #         temp_file.write("%s\n" % item)

    log('FINALIZANDO CRIAÇÃO DE CHAVES CONV115')

    contator_cruza = 0
    ix_115         = -1
    ix_86          = 0
    v_chave_115    = ''
    # Para v_86 existentes em p_chaves_conv86:
    for v_86 in chave_reg_86_ordenada:
    # Enquanto v_chave_115 < v_86[ix_86][0] 

        if (contator_cruza % 10000) == 0:        
            log('[ %s ] NOTAS PROCESSADAS ATÉ O MOMENTO NO CRUZAMENTO DE DADOS.'%contator_cruza)

        while v_chave_115 < v_86[0]:
            # se ix_115 > p_chaves_conv115_.length - 1 
            if  ix_115 > (len(chave_reg_c115_ordenada) - 1):
                v_chave_115 = 'ZZZZ'
            # Senão
            # ix_115 = ix_115 + 1
            # v_chave_115 = v_115[ix_115][0]
            else:
                ix_115 = ix_115 + 1
                v_chave_115 = chave_reg_c115_ordenada[ix_115][0]
        
        # Se v_86[0] > v_chave_115
        if v_chave_115 > v_86[0]:
            # Sem_correspondente(v_86[ix_86][1], lista_86)
            # Sem_correspondente(lista_aprocessar[v_86[1]])            
            lista_sem_correspondentes.append(lista_aprocessar[v_86[1]])

        # Se v_86[0] = v_chave_115
        if v_chave_115 == v_86[0]:
            # Batimento_e_ajuste_valores(v_115[ix_115][1], v_86[ix_86][1], lista_86, lista_115)

            registro_115_mestre = dados_conv115[chave_reg_c115_ordenada[ix_115][1][0]]['mestre'][chave_reg_c115_ordenada[ix_115][1][1]]
            registro_115_item   = dados_conv115[chave_reg_c115_ordenada[ix_115][1][0]]['mestre'][chave_reg_c115_ordenada[ix_115][1][1]]['item'][chave_reg_c115_ordenada[ix_115][1][2]]
            registro            = lista_aprocessar[v_86[1]]

            divergente = False

            if int(registro['VALOR_TOTAL']) != int(registro_115_mestre['VALOR_TOTAL']) or int(registro['BASE_ICMS'])  != int(registro_115_mestre['BASE_ICMS'])\
                                                                                       or int(registro['VALOR_ICMS']) != int(registro_115_mestre['VALOR_ICMS']):                
                divergente = True

            if int(registro['VALOR_ITEM']) != int(registro_115_item['VALOR_TOTAL']):                
                divergente = True
        
            if divergente == True:

                campos_registro_115_mestre = ['NUMERO_NF_115', 'SERIE_115', 'DATA_EMISSAO_115', 'HASH_COD_NF_115', 'CNPJ_CPF_115', 'IE_115', 'RAZAO_SOCIAL_115', 
                                              'CADG_COD_115', 'VALOR_TOTAL_115', 'BASE_ICMS_115', 'VALOR_ICMS_115']
                
                for field in campos_registro_115_mestre:
                    registro[field] = registro_115_mestre[field.replace('_115', '')]

                campos_registro_115_item = ['VALOR_TOTAL_115_ITEM', 'NUM_ITEM_115_ITEM']
                
                for field in campos_registro_115_item:
                    registro[field] = registro_115_item[field.replace('_115_ITEM', '')]

                lista_divergentes.append(registro)

            # else:

            registro['HASH_CODE_ARQ'] = registro_115_mestre['HASH_COD_NF']
            registro['CNPJ_CPF']      = registro_115_mestre['CNPJ_CPF']
            registro['IE']            = registro_115_mestre['IE']
            registro['RAZAOSOCIAL']   = registro_115_mestre['RAZAO_SOCIAL']
            registro['CADG_COD']      = registro_115_mestre['CADG_COD']

            lista_resultado.append(registro)
            contator_cruza += 1

        ix_86 = ix_86 + 1

    log('[ %s ] NOTAS PROCESSADAS.'%contator_cruza)

    retorno_result              = lista_resultado
    retorno_divergencias        = lista_divergentes
    retorno_sem_correspondentes = lista_sem_correspondentes

    log('FINALIZANDO O CRUZAMENTO DE DADOS...')
    
    return [retorno_result,retorno_divergencias,retorno_sem_correspondentes]
        
def geraListaArquivosConv115(uf, ano, mes, serie):

    con = sql.geraCnxBD(configuracoes)

    ano_mes = ano + mes

    query = """SELECT
                     CTR_APUR_DTINI  "Data"
                    ,CTR_SERIE       "Serie"
                    ,CTR_NUM_NFINI   "Inicio_Volume"
                    ,CTR_NUM_NFFIN   "Fim_Volume"
                    ,Volume          "Volume"
                    ,CTR_NF_NOMARQ   "NOME_ARQ_MESTRE"
                    ,CTR_ITEM_NOMARQ "NOME_ARQ_ITEM"
                    ,DIRETORIO       "DIRETORIO"
                FROM(
                    SELECT /*+ parallel(8)*/
                        c.emps_cod
                        ,l.uf_filial
                        ,c.fili_cod
                        ,l.mes_ano
                        ,c.ctr_serie
                        ,lpad(c.ctr_volume, 3, '0') volume
                        ,c.ctr_num_nfini
                        ,c.ctr_num_nffin
                        ,c.ctr_apur_dtini
                        ,c.ctr_nf_nomarq
                        ,c.ctr_item_nomarq
                        ,'/portaloptrib/LEVCV115/' || l.uf_filial || '/' || to_char(c.ctr_apur_dtini, 'YY/MM')
                        || '/TBRA/' || l.fili_cod || '/SERIE/' || l.id_serie_levantamento || '/'
                        || decode(c.ctr_ind_retif, 'N', 'PROTOCOLADO', 'OBRIGACAO') diretorio--SP não é fixo, utilizar '[uf]'
                        ,RANK() OVER(PARTITION BY c.emps_cod, c.fili_cod, c.ctr_serie, c.ctr_volume ORDER BY ctr_ind_retif DESC) seq_retificacao
                      FROM openrisow.ctr_ident_cnv115     c
                      JOIN gfcarga.tsh_serie_levantamento l 
                        ON c.emps_cod = l.emps_cod
                       AND c.fili_cod = l.fili_cod
                       AND to_char(c.ctr_apur_dtini, 'MMYYYY') = to_char(l.mes_ano, 'MMYYYY')
                       AND c.ctr_serie = replace(l.serie, ' ', '')
                     WHERE to_char(l.mes_ano, 'YYYYMM') = '%s'
                       AND l.uf_filial                  = '%s'
                       AND c.CTR_SERIE                 IN ('%s'))
               WHERE SEQ_RETIFICACAO = 1
               ORDER BY ctr_apur_dtini,ctr_serie,ctr_num_nfini        
            """%(ano_mes,uf,serie.strip())
    
    con.executa(query)
    
    return con.fetchall()

if __name__ == "__main__" :
    
    ret = 0
    
    comum.addParametro( 'UF'               , None, "UNIDADE FEDERATIVA"                   , True , '' )
    comum.addParametro( 'DIRETORIO_PLEITO' , None, "DIRETÓRIO DO ARQUIVO A SER PROCESSADO", True , '')

    if not comum.validarParametros() :
        log('### ERRO AO VALIDAR OS PARÂMETROS')
        ret = 91
    else:
        configuracoes.uf         = comum.getParametro('UF')
        configuracoes.dir_pleito = comum.getParametro('DIRETORIO_PLEITO')

        diretorio_arquivos_conv86 = configuracoes.dir_pleito
        
        diretorio_arquivos_conv86 = diretorio_arquivos_conv86.replace('<<UF>>', configuracoes.uf)
        diretorio_arquivos_conv86 = diretorio_arquivos_conv86.replace('<<DIR_PLEITO>>', configuracoes.dir_pleito)
        
        log('DIRETÓRIO.:', diretorio_arquivos_conv86 )    

        if not processar():
            log(' ERRO no processamento! '.center(100,'='))
            ret = 92

    sys.exit(ret)