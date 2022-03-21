#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  MODULO ...: TESHUVA
  SCRIPT ...: loader_1400_municipio.py
  CRIACAO ..: 26/05/2021
  AUTOR ....: VICTOR SANTOS / KYROS TECNOlogIA
  DESCRICAO : 

----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 14/04/2021 - VICTOR SANTOS - Kyros Tecnologia
      - Criacao do script.
        Documentacao : 
    * 02/06/2021 - AIRTON BORGES DA SILVA FILHO - Kyros Tecnologia
      - ALTERADO PARA SP.
        Documentacao : 

        SCRIPT ......: drop_create_SP.py
        AUTOR .......: Victor Santos
        Alteração para novo formato de script
----------------------------------------------------------------------------------------------
"""

import os
import sys

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)

import configuracoes

import datetime
import cx_Oracle
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook

import comum
import sql

log.gerar_log_em_arquivo = True

def dtf():
    return (datetime.datetime.now().strftime('%Y%m%d%H%M%S'))

def DropCreate() :
    con=sql.geraCnxBD(configuracoes)
   
    cmd_sql1="""drop table gfcadastro.reg_1400_municipio_sp"""

    log("Removendo a tabela antiga....")
         
    try : 
        con.executa(cmd_sql1)
        log("TABELA REMOVIDA COM SUCESSO!")
    except Exception as e :
        log('Erro ao executar o Sql ...')
        log(str(e))
        log(cmd_sql1)
        pass
      
    cmd_sql1="""create table gfcadastro.reg_1400_municipio_sp
                (
                    descricao   VARCHAR2(100),
                    uf          VARCHAR2(2),
                    codigo_ibge NUMBER(11),
                    indice      number,
                    codigo_pva  varchar2(30),
                    ano         varchar2(4),
                    mes         varchar2(2),
                    data_inicio date    not null,
                    data_fim    date
                )"""

    log("INICIANDO CRIAÇÃO DA TABELA")
         
    try : 
        con.executa(cmd_sql1)
        log("TABELA CRIADA COM SUCESSO!")
    except Exception as e :
        log('Erro ao executar o Sql ...')
        log(str(e))
        log(cmd_sql1)
        ret = 99
    
    return(0)

if __name__ == "__main__" :
    ret = 0
    txt = ''
    log('Processando ... ')
    comum.carregaConfiguracoes(configuracoes)
    ret = DropCreate()
    if ret != 0:
      log('ERRO NO PROCESSAMENTO')