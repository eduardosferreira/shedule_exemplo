#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
MODULO ...: TESHUVA
SCRIPT ...: retorna_dados_banco.py
CRIACAO ..: 11/01/2022
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
import sql
from openpyxl import Workbook, load_workbook
import openpyxl
import datetime
import shutil
from pathlib import Path
from openpyxl.utils import get_column_letter
import fnmatch

log.gerar_log_em_arquivo = True
comum.carregaConfiguracoes(configuracoes)

def ultimo_Arquivo_Diretorio(p_arq_mascara, p_diretorio):
    v_arquivo = ""
    l_diretorio = Path(p_diretorio)
    l_arq = l_diretorio.glob(p_arq_mascara)
    l_procura_arquivos = sorted(l_arq, reverse=False)
    
    if l_procura_arquivos:        
        for item in l_procura_arquivos:    
            arquivo = os.path.basename(str(item))
            if fnmatch.fnmatch(arquivo,p_arq_mascara) :
                v_arquivo = arquivo
    return v_arquivo

def tipoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r', encoding='iso-8859-1')
        fd.readline()
        fd.close()
    except :
        return 'utf-8'
    return 'iso-8859-1'
    
def converter(uf, data_ini, arquivo):

    dados_retorno = []
    
    v_diretorio = os.path.join(configuracoes.dir_geracao_arquivos, 'PLSQL' )

    log('DIRETÓRIO ...........', v_diretorio )    
       
    v_mascara_arq = arquivo
    v_mascara_arq = v_mascara_arq.replace('<<UF>>', uf)
    v_mascara_arq = v_mascara_arq.replace('<<DATAINI>>', data_ini)
        
    log('MASCARA .............', v_mascara_arq)

    arquivo = ultimo_Arquivo_Diretorio(v_mascara_arq, v_diretorio)
         
    log('ARQUIVO .............', arquivo)
    if arquivo:
        log('Lendo arquivo ', arquivo)
        fd = open(os.path.join(v_diretorio,arquivo), 'r', encoding=tipoArquivo(os.path.join(v_diretorio,arquivo)) )

        cabecalho = fd.readline().replace('"', '').replace('\n', '').split(';')
        
        for l in fd:

            dic_linha = {}

            if l.endswith('\n'):
                l = l[:-1]
            
            if l.endswith(';'):
                l = l[:-1]
            
            linha = l.split('";"')
            
            for i in range(len(cabecalho)):

                dic_linha[cabecalho[i]] = linha[i].replace('"', '').replace('\n', '')
            
            dados_retorno.append(dic_linha)

        fd.close()

    return { 'dados' : dados_retorno,  'status' : 'Ok' }