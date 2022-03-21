#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
MODULO ...: TESHUVA
SCRIPT ...: retorna_dados_banco.py
CRIACAO ..: 05/01/2022
AUTOR ....: Victor Santos / Welber Pena - KYROS TECNOLOGIA
DESCRICAO.: 
----------------------------------------------------------------------------------------------
  HISTORICO : 
----------------------------------------------------------------------------------------------
"""
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
from pathlib import Path
from openpyxl.utils import get_column_letter

log.gerar_log_em_arquivo = True
comum.carregaConfiguracoes(configuracoes)


def valida_sql(sql):
    carac_invalidos = [ 'DECLARE', 'BEGIN', 'END', 'TRUNCATE', 'CREATE', 'DROP', 'GRANT', 'UPDATE', 'DELETE', 'INSERT' ]
    sql_ok = True
    for c in sql :
        if c in carac_invalidos :
            sql_ok = False
    return sql_ok


def le_arquivo_sql(arquivo_sql):
    dic_retorno = { 'status' : 'ERRO' }
    query_exec = """"""
    for row in open(arquivo_sql, 'r').readlines():
        linha = row.replace('\n', '').replace('\r', '')
        if linha:
            query_exec += linha + '\n'

    query      = query_exec.upper()
   
    valida     = valida_sql(query)

    if valida == False:
        log('#### ERRO, SQL INVALIDO, INSTRUÇÕES QUE COMEÇAM COM DECLARE, BEGIN, END, TRUNCATE, CREATE, DROP, GRANT, -, ;, / NÃO SÃO PERMITIDAS.')
        return dic_retorno
    
    return { 'sql' : query_exec, 'status' : 'Ok' }


def converter(query_exec):
    con=sql.geraCnxBD(configuracoes)
    dic_retorno = { 'status' : 'ERRO' }

    query      = query_exec.upper()

    if query.startswith('SELECT') or query.startswith('WITH'):
            
        if not query.__contains__('WHERE'):
            log('#### ERRO, CONFIRA SUA QUERY, NÃO FOI ENCONTRADA A CLÁUSULA WHERE E ISTO GERA RISCO AOS DADOS. CASO QUEIRA EXECUTAR, ADICIONE WHERE 1 = 1') #mudar frase
            return dic_retorno

        log('Executando QUERY, aguarde...')
        log(' QUERY '.center(80,'*'), '\n', query_exec, '\n', '*'*80)
        retorno   = []
        cabecalho = []
        tipo      = []
        con.executa(query_exec)
        colunas = con.description()
        result  = con.fetchone()
        count   = 0
        if not result:
            log("#### ATENÇÃO: Nenhum Resultado para query")
            for col in colunas:
                cabecalho.append(col[0])
        else:                    
            try:     
                for col in colunas:
                    cabecalho.append(col[0])
                    tipo.append(str(col[1]).split('.')[1].split("'")[0])
                while result:
                    registro = {}
                    for i in range(len(cabecalho)):
                        nome_col = cabecalho[i]
                        registro[nome_col] = result[i]

                    retorno.append(registro) 
                    result = con.fetchone()
                    count += 1

                log('Quantidade de linhas para esta query: ', count)               
            except Exception as e:
                log('ERRO - ' , e)
                return dic_retorno

    else:
        log('#### ERRO, CONFIRA SUA QUERY, NÃO FOI ENCONTRADA A CLÁUSULA ( SELECT OU WITH ) NO INÍCIO DA INSTRUÇÃO.')
        return dic_retorno
    
    return { 'dados' : retorno, 'tipo_colunas' : tipo, 'cabecalho': cabecalho,  'status' : 'Ok' }