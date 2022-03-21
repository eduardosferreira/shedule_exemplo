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
import os
import sys

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)

import configuracoes
import comum
import layout

log.gerar_log_em_arquivo = True
comum.carregaConfiguracoes(configuracoes)
layout.carregaLayout()

def tipoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r', encoding='utf-8')
        fd.readline()
        fd.close()
    except :
        return 'iso-8859-1'
    return 'utf-8'

def converte_conv86_to_json( path_arquivo ):

    log('INICIANDO A LEITURA DO ARQUIVO CONV86')
    log('PATH', path_arquivo)

    if not os.path.isfile(path_arquivo):
        
        log('ERRO, NÃO EXISTE ESSE ARQUIVO NO DIRETÓRIO', path_arquivo)
        return []
    else:        
        arquivo = open(path_arquivo, 'r', encoding=tipoArquivo(os.path.join(path_arquivo)))
        log('ENCONTRADO ARQUIVO NO DIRETORIO, INICIANDO QUEBRA DO REGISTRO')

    lista_dados = []

    numero_da_linha = 0

    for registro in arquivo:

        numero_da_linha += 1
        
        if registro.startswith('1'):
            
            registro_quebrado = layout.quebraRegistroDicionario(registro, 'LayoutConv86_1')
            item = []
            registro_quebrado['item']         = item
            registro_quebrado['numero_linha'] = numero_da_linha
            lista_dados.append(registro_quebrado)

        else:
            registro_quebrado = layout.quebraRegistroDicionario(registro, 'LayoutConv86_2')
            registro_quebrado['numero_linha'] = numero_da_linha
            item.append(registro_quebrado)

    log('ARQUIVO CONV86 RETORNADO COM SUCESSO')
    log('TOTAL DE LINHAS ...: {:,d}'.format(numero_da_linha).replace(',','.'))
    return lista_dados

if __name__ == "__main__" :
    
    ret = 0
    
    comum.addParametro( 'DIRETORIO_ARQUIVOS', None, "diretorio_arquivos" , True , '' )

    if not comum.validarParametros() :
        log('### ERRO AO VALIDAR OS PARÂMETROS')
        ret = 91
    else:
        configuracoes.diretorio_arquivos  = comum.getParametro('DIRETORIO_ARQUIVOS')

        if not converte_conv86_to_json(configuracoes.diretorio_arquivos):
            log('ERRO NO PROCESSAMENTO!')
            ret = 92

    sys.exit(ret)
