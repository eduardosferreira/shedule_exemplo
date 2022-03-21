#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: GF
  MODULO ...:
  SCRIPT ...: controlador_valida_PVA.py
  CRIACAO ..: 02/03/2022
  AUTOR ....: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO:
    
----------------------------------------------------------------------------------------------
  HISTORICO:
    * 02/03/2022 - Welber Pena de Sousa - Kyros Tecnologia
        - Criacao do script.
----------------------------------------------------------------------------------------------
    
----------------------------------------------------------------------------------------------
"""

from datetime import datetime
import multiprocessing
import sys
import os
import shutil
import calendar
import glob
import threading
import cx_Oracle
import time

SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)

import configuracoes
import comum
import sql
from controlador_por_id_serie import *

log.gerar_log_em_arquivo = True

status_thread = {}

def processar():
    mes_ano_inicial = comum.getParametro('MES_ANO_INICIAL')
    mes_ano_final   = comum.getParametro('MES_ANO_FINAL')
    uf              = comum.getParametro('UF')
    id_serie        = comum.getParametro('ID_SERIE')
    series          = comum.getParametro('SERIES')
    series_not_in   = comum.getParametro('SERIES_NOT_IN')
    id_execucao     = int(comum.getParametro('ID_EXECUCAO'))
    usuario         = comum.getParametro('USUARIO')
    v_max_threads   = int(comum.getParametro('MAX_THREADS'))
    nome_script     = configuracoes.script

    ### Parametro opcional
    # p_processar     = comum.getParametro('PROCESSAR')
    
    if not os.path.isfile( nome_script ) :
        log('ERRO - script configurado no .cfg << %s >> nao encontrado no diretorio'%(nome_script))

    if id_serie :
        v_lista_id_serie_levantamento = [[id_serie,],]
    else :
        v_lista_id_serie_levantamento = retornaIdSerieExecutar( mes_ano_inicial, mes_ano_final, uf, series, series_not_in )

    v_qtde_erros = 0
    thread = 0

    if v_lista_id_serie_levantamento :
        print(dir(configuracoes))
        v_lista_threads = []
        
        while v_lista_id_serie_levantamento :
            if len(v_lista_threads) < v_max_threads :
                id_serie_levantamento = v_lista_id_serie_levantamento.pop(0)[0]
                v_parametros = "%s"%( id_serie_levantamento )
                id_sparta = criaExecucao(id_execucao, id_serie_levantamento, nome_script, v_parametros, usuario ) 
                                         
                #multiprocessing inicia a execução do script valida_PVA.py (subprocesso) em uma nova thread .
                thread += 1
                th = threading.Thread( target= executaProcesso, args = [id_sparta, nome_script, v_parametros, thread] )
                th.start()
                time.sleep(2)
                v_lista_threads.append([ th, id_sparta, thread ])
            else :

                v_lista_threads, erros = retiraThreadsFinalizadas( v_lista_threads, status_thread ) 
                v_qtde_erros += erros
                time.sleep(30)
        
        while v_lista_threads : 
            v_lista_threads, erros = retiraThreadsFinalizadas( v_lista_threads, status_thread ) 
            v_qtde_erros += erros
            time.sleep(30)


    ## Apos todas as Threads finalizadas.
    if v_qtde_erros > 0 :
        log('ERRO - Execuções com ERRO .. Verifique ...')
        return False

    log('Execuções com SUCESSO !!!')

    log('-'*150) 
    
    return True 


def executaProcesso(id_sparta, nome_script, v_parametros, idx_thread ) :
    global status_thread
    status_thread[idx_thread] = 1
    comando = "./%s %s"%( nome_script, v_parametros )
    log('Executando comando : %s'%(comando))
    path_log = '%s/%s_%s_%s.log'%(configuracoes.dir_log, nome_script, id_sparta, datetime.now().strftime('%Y%m%d%H%M%S'))
    comando += ' > %s'%(path_log)
    os.system(comando)
    
    if os.path.isfile( path_log ) :
        fd = open(path_log, 'r')
        status = False
        msg_erro = False
        for l in fd :
            if l.__contains__('STATUS da execucao :') :
                status = l.replace('\n','').split(' : ')[-1].strip()
            elif l.upper().__contains__('### ERRO') :
                msg_erro = "ERRO"
                msg_erro += l.replace('\n','').split('### ERRO')[-1]
        if status == 'SUCESSO' :
            status_thread[idx_thread] = 0
            return 0
        fd.seek(0,0)
        txt = """%s
"""%( '  LOG da execucao com ERRO  '.center(120, '=') )
        for l in fd :
            txt += l.replace('\n','')
            txt += '\n'
        
        txt += """
%s"""%( '  FIM do LOG da execucao com ERRO  '.center(120, '-') )
        log(txt)
    status_thread[idx_thread] = msg_erro if msg_erro else 2
    return 2


if __name__ == "__main__":
    ret = 0
    comum.carregaConfiguracoes(configuracoes)

    if not os.path.isfile(configuracoes.script) :
        log('ERRO - Nao existe o script :', configuracoes.script)
        ret = 1
    
    if not ret :
        log(str("  INICIANDO CONTROLADOR POR SERIE - %s  "%(configuracoes.script)).center(120,'#'))
        comum.carregaConfiguracoes(configuracoes)
        
        ### Parametros padroes, todos os controladoresSPT devem ter :
        comum.addParametro('ID_EXECUCAO',     '-ID', "Id_execucao do Sparta.", True)
        comum.addParametro('MES_ANO_INICIAL', '-MI', "Mes e ano inicial no formato YYYYmm.", True)
        comum.addParametro('MES_ANO_FINAL',   '-MF', "Mes e ano final no formato YYYYmm.", True)
        comum.addParametro('UF',              '-UF', "UF dos dados a processar.", False)
        comum.addParametro('ID_SERIE',        '-IS', "ID_SERIE_LEVANTAMENTO a processar.", False, False, False)
        comum.addParametro('SERIES',          '-S',  "Series a processar.", False)
        comum.addParametro('SERIES_NOT_IN',   '-SN', "Series que serao desconsideradas (NOT IN) .", False)
        comum.addParametro('USUARIO',         '-U',  "Nome do usuario que iniciou o processo no Sparta .", False)
        comum.addParametro('MAX_THREADS',     '-MT', 'Maximo de threads executadas em paralelo ( 1 ate 5 )', True, '2', 5)

        ### Parametros opcionais, (parametro do script que sera chamado.)
        # comum.addParametro('PROCESSAR',       '-P',  'Dados a processar.( Mestre, Item, Controle, Destinatario, Todos )', True, 'Todos', 'Todos')
    
        if not comum.validarParametros():
            ret = 2
        else:
            try :
                if not processar():
                    ret = 3
            except Exception as e :
                log('ERRO ao processar:', e)
                ret = 99
            # if (ret > 0):
            log("### Retorno da execução ..:", ret)

        log(str("  FIM DO CONTROLADOR POR SERIE - %s  "%(configuracoes.script)).center(120,'#'))
    
    sys.exit(ret)


