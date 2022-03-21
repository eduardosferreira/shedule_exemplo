#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-

"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: Sparta
  SCRIPT ...: comparacao_conv86.py
  CRIACAO ..: 15/02/2022
  AUTOR ....: Welber Pena de Sousa / Kyros Consultoria
  DESCRICAO : PTITES-1577
                - Documentacao no link abaixo
                    https://wikicorp.telefonica.com.br/display/PTI/%5BPTITES-1568%5D+-+%5BPTITES-1577%5D++-+RS06+Python+-+comparacao_conv86.py+-+novo

----------------------------------------------------------------------------------------------
  HISTORICO :
    * 15/02/2022 - Welber Pena de Sousa - Kyros Consultoria 
        - Criacao do script.
    
----------------------------------------------------------------------------------------------
"""

import os
import sys
import fnmatch

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)

import configuracoes
import comum
import layout
import conversor.converte_conv86_to_json as converte_conv86_to_json


log.gerar_log_em_arquivo = True


def processar() :
    mascara = '*.txt'
    log('  Iniciando processamento  '.center(100, '-'))
    diretorio_arquivos_conv86 = os.path.join(configuracoes.dir_entrada, configuracoes.uf, configuracoes.dir_pleito)
    log('Processando diretorio : %s'%(diretorio_arquivos_conv86))    
    arquivo_conv86 = selecionaArquivoTXT(diretorio_arquivos_conv86)
    
    if not arquivo_conv86 :
        log('ERRO - Não encontrado arquivo com a mascara %s \n no diretorio %s'%( mascara, diretorio_arquivos_conv86 ))
        return False
    
    log('Encontrado arquivo : %s'%(arquivo_conv86))
    path_arquivo_conv86 = os.path.join(diretorio_arquivos_conv86, arquivo_conv86)

###############################  ORIGINAL  ########################################
    dados_arquivo_KRONOS = converte_conv86_to_json.converte_conv86_to_json( path_arquivo_conv86 )

    ## Ordena a lista de itens pelos campos : DATA_EMISSAO, SERIE, NUMERO_NF, NUM_ITEM    
    log('Iniciando ordenação do arquivo original')
    dados_tipo_2_KRONOS = comum.ordenaListaDicionarios( dados_arquivo_KRONOS[0]['item'], ['SERIE', 'NUMERO_NF', 'DATA_EMISSAO',  'NUM_ITEM'] )
    lista_chaves_2_KRONOS = {}
    idx = 0
    for registro in dados_tipo_2_KRONOS :
        chave = "%s|%s|%s"%( registro['SERIE'], registro['NUMERO_NF'], registro['DATA_EMISSAO'] )
        if chave not in lista_chaves_2_KRONOS :
            lista_chaves_2_KRONOS[chave] = idx
        idx += 1
    log('Finalizando ordenação do arquivo original')
    
    ### Definir onde ficara o arquivo a ser comparado .. ????
    # diretorio_arquivo_TESHUVA = os.path.join(configuracoes.dir_geracao_arquivos, configuracoes.uf, configuracoes.dir_pleito_teshuva)
    diretorio_arquivo_TESHUVA = os.path.join(configuracoes.dir_entrada, configuracoes.uf, configuracoes.dir_pleito_teshuva).replace('DEV', 'PROD')
    log('Processando diretorio comparacao : %s'%(diretorio_arquivo_TESHUVA))
    arquivo_TESHUVA = selecionaArquivoTXT(diretorio_arquivo_TESHUVA)
    if not arquivo_TESHUVA : 
        log('ERRO - Não encontrado arquivo de comparacao com a mascara %s \n no diretorio %s'%( mascara, diretorio_arquivo_TESHUVA ))
        return False
    log('Encontrado arquivo comparacao : %s'%(arquivo_TESHUVA))
    path_arquivo_TESHUVA = os.path.join(diretorio_arquivo_TESHUVA, arquivo_TESHUVA)
    dados_TESHUVA_conv86 = converte_conv86_to_json.converte_conv86_to_json( path_arquivo_TESHUVA )

    ## Ordena a lista de itens pelos campos : DATA_EMISSAO, SERIE, NUMERO_NF, NUM_ITEM
    log('Iniciando ordenação do arquivo TESHUVA')
    dados_tipo_2_TESHUVA = comum.ordenaListaDicionarios( dados_TESHUVA_conv86[0]['item'], ['SERIE', 'NUMERO_NF', 'DATA_EMISSAO',  'NUM_ITEM'] )
    lista_chaves_2_TESHUVA = {}
    idx = 0
    for registro in dados_tipo_2_TESHUVA :
        chave = "%s|%s|%s"%( registro['SERIE'], registro['NUMERO_NF'], registro['DATA_EMISSAO'] )
        if chave not in lista_chaves_2_TESHUVA :
            lista_chaves_2_TESHUVA[chave] = idx
        idx += 1
    log('Finalizando ordenação do arquivo TESHUVA')

###############################  PARA TESTES  ########################################
    

    # ## Ordena a lista de itens pelos campos : DATA_EMISSAO, SERIE, NUMERO_NF, NUM_ITEM    
    # log('Iniciando ordenação do arquivo original')
    # fd = open('dados_tipo_2_KRONOST.txt', 'r' )
    # l = fd.readlines()
    # dados_tipo_2_KRONOS = eval(l[0])
    # fd.close()
    # log('Finalizando ordenação do arquivo original')
    
    # ### Definir onde ficara o arquivo a ser comparado .. ????
    # diretorio_arquivo_TESHUVA = os.path.join(configuracoes.dir_geracao_arquivos, configuracoes.uf, configuracoes.dir_pleito_teshuva)
    # log('Processando diretorio comparacao : %s'%(diretorio_arquivo_TESHUVA))
    # arquivo_TESHUVA = selecionaArquivoTXT(diretorio_arquivo_TESHUVA)
    # if not arquivo_TESHUVA : 
    #     log('ERRO - Não encontrado arquivo de comparacao com a mascara %s \n no diretorio %s'%( mascara, diretorio_arquivo_TESHUVA ))
    #     return False
    # log('Encontrado arquivo comparacao : %s'%(arquivo_TESHUVA))
    # path_arquivo_TESHUVA = os.path.join(diretorio_arquivo_TESHUVA, arquivo_TESHUVA)

    # ## Ordena a lista de itens pelos campos : DATA_EMISSAO, SERIE, NUMERO_NF, NUM_ITEM
    # log('Iniciando ordenação do arquivo TESHUVA')
    # fd = open('dados_tipo_2_TESHUVA.txt', 'r' )
    # l = fd.readlines()
    # dados_tipo_2_TESHUVA = eval(l[0])
    # fd.close()
    # log('Finalizando ordenação do arquivo TESHUVA')

#######################################################################




    # resultado_divergencias = []
    # resultado_sem_correspondentes = []

    resultado_divergencias, resultado_sem_correspondentes = cruza_dados( dados_tipo_2_KRONOS, dados_tipo_2_TESHUVA, lista_chaves_2_TESHUVA )
    
    # resultado_divergencias += divergencias
    # resultado_sem_correspondentes += sem_correspondentes
    log( 'Encontrado %s registros com divergencias.'%len(resultado_divergencias) )
    log( 'Encontrado %s registros sem correspondentes.'%len(resultado_sem_correspondentes) )

    divergencias, sem_correspondentes = cruza_dados( dados_tipo_2_TESHUVA, dados_tipo_2_KRONOS, lista_chaves_2_KRONOS)

    log( 'Encontrado %s registros com divergencias.'%len(divergencias) )
    log( 'Encontrado %s registros sem correspondentes.'%len(sem_correspondentes) )
    log('Somando lista de divergente e sem correspondentes')
    # resultado_divergencias += divergencias
    resultado_sem_correspondentes += sem_correspondentes
    
    ### 2 – Gerar o arquivo divergencias :
    if resultado_divergencias :
        log('  Gerando relatorio de divergencias  '.center(100, '='))
        # print(resultado_divergencias[0].keys())
        ### Campos do registro e do registro_comparado, intercalados ??????????
        ### Exemplo : TIPO;TIPO_comaprado;MODELO;MODELO_comparado;NUMERO_NF;NUMERO_NF_comparado ...
        
        diretorio_arquivo_divergencias = os.path.join( configuracoes.dir_geracao_arquivos, configuracoes.uf, configuracoes.dir_pleito )
        nome_arquivo_divergencias = arquivo_conv86.split('.')[0] + '_DIVERGENCIAS.csv'
        log('Diretorio de geracao ..: %s'%( diretorio_arquivo_divergencias ))
        log('Arquivo gerado ........: %s'%( nome_arquivo_divergencias ))
        log('Qtde. de linhas .......: {:d}'.format( len(resultado_divergencias) ).replace(',', '.')) 
        campos_divergencias = resultado_divergencias[0].keys()
        arquivo_divergencia = open( os.path.join( diretorio_arquivo_divergencias, nome_arquivo_divergencias ), 'w' )
        arquivo_divergencia.write(';'.join( x for x in campos_divergencias ))
        arquivo_divergencia.write('\n')
        resultado = comum.ordenaListaDicionarios( resultado_divergencias, ['numero_linha'] )
        for registro in resultado :
            linha = [ str(registro[campo]) for campo in campos_divergencias ]
            arquivo_divergencia.write(';'.join( x for x in linha ))
            arquivo_divergencia.write('\n')

        arquivo_divergencia.close()
        log('  Fim da geracao do relatorio de divergencias  '.center(100, '-'))

    ### 3 – Gerar o arquivo sem correspondentes:
    if resultado_sem_correspondentes :
        log('  Gerando relatorio de sem correspondentes  '.center(100, '='))
        # print(resultado_sem_correspondentes[0].keys())
        diretorio_arquivo_sem_correspondentes = os.path.join( configuracoes.dir_geracao_arquivos, configuracoes.uf, configuracoes.dir_pleito )
        if not os.path.isdir( diretorio_arquivo_sem_correspondentes ) :
            os.makedirs(diretorio_arquivo_sem_correspondentes)
        nome_arquivo_sem_correspondentes = arquivo_conv86.split('.')[0] + '_SEM_CORRESPONDENTE.csv'
        log('Diretorio de geracao ..: %s'%( diretorio_arquivo_sem_correspondentes ))
        log('Arquivo gerado ........: %s'%( nome_arquivo_sem_correspondentes ))
        log('Qtde. de linhas .......: {:d}'.format( len(resultado_sem_correspondentes) ).replace(',', '.'))

        campos_sem_correspondentes = [ 'TIPO', 'MODELO', 'NUMERO_NF', 'SERIE', 'DATA_EMISSAO', 'HASH_CODE_ARQ', 'CNPJ_CPF', 'IE', 'RAZAOSOCIAL', 'CADG_COD', 'VALOR_TOTAL', 'BASE_ICMS', 
                                    'VALOR_ICMS', 'NUM_ITEM', 'VALOR_ITEM', 'ICMS_ESTORNO', 'HIPOTESE_ESTORNO', 'MOTIVO_ESTORNO', 'NUM_RECLAMACAO' ]
        arquivo_sem_correspondente = open( os.path.join( diretorio_arquivo_sem_correspondentes, nome_arquivo_sem_correspondentes ), 'w' )
        arquivo_sem_correspondente.write(';'.join( x for x in campos_sem_correspondentes ))
        arquivo_sem_correspondente.write('\n')
        resultado = comum.ordenaListaDicionarios( resultado_sem_correspondentes, ['numero_linha'] )
        for registro in resultado :
            linha = [ registro[campo] for campo in campos_sem_correspondentes ] 
            arquivo_sem_correspondente.write(';'.join( x for x in linha ))
            arquivo_sem_correspondente.write('\n')

        arquivo_sem_correspondente.close()
        log('  Fim da geracao do relatorio de sem correspondentes  '.center(100, '-'))

    return True ### Fim da subrotina.


def cruza_dados( dados_tipo_2_entrada, dados_tipo_2_comparar, lista_chaves_2_comparar ) :
    resultado_divergencias = []
    resultado_sem_correspondentes = []

    indice_dados_tipo_2_entrada = 0
    indice_dados_tipo_2_comparar = 0
    indice_ultimo_encontrado = 0
    # print('>>>>', dados_tipo_2_entrada[indice_dados_tipo_2_entrada])
    # print(dados_tipo_2_entrada[indice_dados_tipo_2_entrada].keys())
    lista_campos = [ 'MODELO', 'NUMERO_NF', 'SERIE', 'DATA_EMISSAO', 'HASH_CODE_ARQ', 'CNPJ_CPF', 'IE', 'RAZAOSOCIAL', 'CADG_COD',
                     'VALOR_TOTAL', 'BASE_ICMS', 'VALOR_ICMS', 'NUM_ITEM', 'VALOR_ITEM', 'ICMS_ESTORNO', 
                     'HIPOTESE_ESTORNO', 'MOTIVO_ESTORNO', 'NUM_RECLAMACAO']
    log('  Comparando arquivos ... Aguarde ...  '.center(100, '-'))
    qtde_reg_total = len(dados_tipo_2_entrada)
    while indice_dados_tipo_2_entrada < len(dados_tipo_2_entrada) :
    # while indice_dados_tipo_2_entrada < 20 :
        registro = dados_tipo_2_entrada[indice_dados_tipo_2_entrada]
        if indice_dados_tipo_2_entrada % 100000 == 0 :
            log('Processadas', indice_dados_tipo_2_entrada, 'notas de', qtde_reg_total, 'notas.')
        
        encontrou = False
        encontrou_igual = False
        # log('- Procurando nota ......:', registro['NUMERO_NF'])
        # log('> Dados nota .......:', registro['SERIE'], registro['NUMERO_NF'], registro['DATA_EMISSAO'], registro['NUM_ITEM'])
        # log('> Dados a comparar .:', indice_dados_tipo_2_comparar, registro_comparar['SERIE'], registro_comparar['NUMERO_NF'], registro_comparar['DATA_EMISSAO'], registro_comparar['NUM_ITEM']) 

        chave = "%s|%s|%s"%( registro['SERIE'], registro['NUMERO_NF'], registro['DATA_EMISSAO'] )
        indice_dados_tipo_2_comparar = lista_chaves_2_comparar.get(chave, indice_ultimo_encontrado)-1
        registro_comparar = dados_tipo_2_comparar[indice_dados_tipo_2_comparar]

        while encontrou == False and indice_dados_tipo_2_comparar < len(dados_tipo_2_comparar) :
            # print('SERIES', registro['SERIE'], registro_comparar['SERIE'])
            if registro['SERIE'] != registro_comparar['SERIE'] :
                indice_dados_tipo_2_comparar += 1
                if indice_dados_tipo_2_comparar < len(dados_tipo_2_comparar) :
                    registro_comparar = dados_tipo_2_comparar[indice_dados_tipo_2_comparar]
                continue
            
            # print('DATA_EMISSAO', registro['DATA_EMISSAO'] , registro_comparar['DATA_EMISSAO'])
            if int(registro['DATA_EMISSAO']) != int(registro_comparar['DATA_EMISSAO']) :
                indice_dados_tipo_2_comparar += 1
                if indice_dados_tipo_2_comparar < len(dados_tipo_2_comparar) :
                    registro_comparar = dados_tipo_2_comparar[indice_dados_tipo_2_comparar]
                continue
            # else :
            #     indice_dados_tipo_2_comparar = len(dados_tipo_2_comparar)
            
            if indice_dados_tipo_2_comparar < len(dados_tipo_2_comparar) :
                registro_comparar = dados_tipo_2_comparar[indice_dados_tipo_2_comparar]
                # if ( indice_dados_tipo_2_comparar - indice_ultimo_encontrado )  < 10 :
                    # log('> Dados a comparar .:', indice_dados_tipo_2_comparar, registro_comparar['SERIE'], registro_comparar['NUMERO_NF'], registro_comparar['DATA_EMISSAO'], registro_comparar['NUM_ITEM']) 

                if registro['NUMERO_NF'] == registro_comparar['NUMERO_NF'] and registro['NUM_ITEM'] == registro_comparar['NUM_ITEM'] :
                    encontrou = True
                    encontrou_igual = True
                    for campo in lista_campos :
                        if registro[campo] != registro_comparar[campo] :
                            encontrou_igual = False

                    if not encontrou_igual :
                        for campo in lista_campos :
                            registro[campo+'_comparado'] = registro_comparar[campo]

                else :
                    indice_dados_tipo_2_comparar += 1
                    if indice_dados_tipo_2_comparar < len(dados_tipo_2_comparar) :
                        registro_comparar = dados_tipo_2_comparar[indice_dados_tipo_2_comparar]
                    if ( indice_dados_tipo_2_comparar - indice_ultimo_encontrado )  > 100 :
                        break

        if not encontrou :
            # log('> NAO encontrada !')
            # print(registro)
            resultado_sem_correspondentes.append(registro)
        else :
            indice_ultimo_encontrado = indice_dados_tipo_2_comparar
            if not encontrou_igual :
                resultado_divergencias.append(registro)

        if encontrou_igual :
            # log('> NF Identica.')
            indice_ultimo_encontrado = indice_dados_tipo_2_comparar
        else :
            indice_dados_tipo_2_comparar = indice_ultimo_encontrado

        indice_dados_tipo_2_entrada += 1
        # print(indice_dados_tipo_2_entrada)

    log('  Fim da comparação dos arquivos.  '.center(100, '-'))

    return resultado_divergencias, resultado_sem_correspondentes


def selecionaArquivoTXT( diretorio, mascara = '*.txt' ) :
    arquivo = False
    for arq in os.listdir(diretorio) :
        if fnmatch.fnmatch(arq, mascara) :
            if os.path.isfile(os.path.join(diretorio, arq)) :
                arquivo = arq

    if not arquivo and mascara != '*.zip' :
        arq_zip = selecionaArquivoTXT( diretorio, '*.zip' )
        if arq_zip :
            os.system('cd %s ; unzip %s ; cd - '%( diretorio, arq_zip ))
            arquivo = selecionaArquivoTXT( diretorio, mascara )

    return arquivo


if __name__ == "__main__":
    ret = 0
    
    comum.addParametro( 'UF'                , None, "UF dos dados a processar."                        , True , 'SP' )
    comum.addParametro( 'DIR_ENT_PLEITO'    , None, "Nome do diretorio pleito de entrada ."            , True , 'pleito_2017_regeracao_2020')
    comum.addParametro( 'DIR_PLEITO_TESHUVA', None, "Nome do diretorio pleito de geracao do TESHUVA."  , True , 'pleito_2017_regeracao_2022')

    configuracoes.dir_entrada = ['/'] + configuracoes.dir_entrada.split('/')[:-1]
    configuracoes.dir_entrada.append('conv86')
    configuracoes.dir_entrada = os.path.join( *configuracoes.dir_entrada )

    if not comum.validarParametros() :
        log('### ERRO AO VALIDAR OS PARÂMETROS')
        ret = 91
    else:
        configuracoes.uf                  = comum.getParametro('UF')
        configuracoes.dir_pleito          = comum.getParametro('DIR_ENT_PLEITO')
        configuracoes.dir_pleito_teshuva  = comum.getParametro('DIR_PLEITO_TESHUVA')
        layout.carregaLayout()
        if not processar() :
            log('ERRO durante o processamento !... Finalizando ...')
            ret = 92
        
    sys.exit(ret)