#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: cargaProtocolado.py
  CRIACAO ..: 18/02/2022
  AUTOR ....: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO : Processo para realizar a carga dos arquivos de Conv86 .
                Parametros :
                    - UF 
                    - Periodo
                    - Serie ( Opcional )
                a partir da UF, período e Série (opcional, caso não informado considera todas) 
                informado identifica o diretório protocolado (usando tabela tsh_serie_levantamento) 
                e realiza a execução das linhas de comando para o respectivo CTL (loader - sqlldr)

----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 18/02/2022 - Welber Pena de Sousa - Kyros Tecnologia
        - Criacao do script.

   
----------------------------------------------------------------------------------------------
"""

from fnmatch import fnmatch
import os
import sys

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
dir_execucao = os.path.realpath('.')
sys.path.append(dir_base)
import configuracoes

import time
import cx_Oracle
import datetime
import calendar


import threading
import traceback

import comum
import sql
import layout
import util

log.gerar_log_em_arquivo = True

idControle = ''
codigoUF = ''
LayoutDestinatario = ''
LayoutMestre   = ''
dic_registros = {}
dic_layouts = {}
dic_campos = {}
variaveis = {}
dic_fd = {}
listaMestreProtocolado = []

def processar():     
    global listaMestreProtocolado

    log("Iniciando carga do(s) arquivo(s) de Conv86 ...")
    
    uf = variaveis['UF']
    versao = variaveis['VERSAO']
    registros_processados = 0

    diretorio_arquivos_conv86 = configuracoes.dir_entrada.replace('DEV','PROD').split('/')[:-1]
    diretorio_arquivos_conv86 = os.path.join( '/', *diretorio_arquivos_conv86, 'conv86', uf, versao )
    log('Diretorios de arquivos Conv86 .:', diretorio_arquivos_conv86)
    arquivos = []

    if not os.path.isdir(diretorio_arquivos_conv86) :
        log('ERRO Diretorio Conv86 NAO EXISTE ... Finalizando !!!')
        log('Diretorio ..: %s'%(diretorio_arquivos_conv86))
        ret = 1
    else :
        log('Iniciando processamento ...')
        registros_processados             = len(arquivos)
        for arq in  os.listdir(diretorio_arquivos_conv86) :
            if fnmatch(arq,'*.txt') :
                arquivos.append(arq)

    if arquivos :
        log("# Carregando arquivos ...")
        log("- diretorio de arquivos :", diretorio_arquivos_conv86)

        lst_threads = []

        th = threading.Thread( target = carga, args = ( diretorio_arquivos_conv86, arquivos, os.getpid(), uf) )
        th.start()
        lst_threads.append(th)
        time.sleep(2)

        while threading.active_count() > 1 :
            time.sleep(10)
        
    log("-------------------------------------------------------------------------")
    log('*', " R E S U M O ".center(69,'-') , '*')
    log("* registros_processados ...............:", str(registros_processados))
    log("-------------------------------------------------------------------------")

    log("* Processamento Finalizado !")
    log("-------------------------------------------------------------------------")
    retorno = 0 if not os.path.isfile( 'ERRO_%s'%(os.getpid())) else 5
    if retorno :
        os.remove('ERRO_%s'%(os.getpid()))

    return retorno


def carga(diretorio, lst_arquivos, pid, puf='SP' ):

    os.chdir( log.path_log )

    modelo_ctl = 'CNV86_TIPO2.ctl' if puf == 'SP' else 'CNV86_TIPO2_OUTROS.ctl'
    arq_ctl = '%s_%s'%(  pid, modelo_ctl )
    if not os.path.isfile( os.path.join(dir_execucao, modelo_ctl) ) :
        log('Erro falta arquivo de modelo (.ctl) : ', modelo_ctl)
        return False 

    retorno = True
    for arq in lst_arquivos :
        ##### Monta o novo arquivo .ctl
        fd = open(os.path.join(dir_execucao,modelo_ctl), 'r')
        fd_ctl = open('./%s'%(arq_ctl), 'w')
        # vol = int(arq.split('.')[-1])
        for lin in fd :
            if lin.__contains__('$$UF$$') :
                lin = lin.replace('$$UF$$', variaveis['UF'] )
            
            elif lin.__contains__('$$VER$$') :
                lin = lin.replace('$$VER$$', variaveis['VERSAO'] )
            
            elif lin.__contains__('$$ARQ$$') :
                lin = lin.replace('$$ARQ$$', arq )
            
            elif lin.__contains__('$$IN$$') :
                lin = lin.replace('$$IN$$', os.path.join( diretorio, arq ) )

            fd_ctl.write(lin)

        fd.close()
        fd_ctl.close()
        
        ## Executa o loader.
        # comando = """sqlldr openrisow/openrisow@%s %s data=\'%s\' > log/%s.log"""%( variaveis['banco'], os.path.join( os.path.realpath('.'), 'ctl', arq_ctl ), os.path.join(diretorio, arq.replace(' ', '\\ ')), arq_ctl[:-4] )
        # comando = """sqlldr gfcarga/vivo2019@%s %s > log/%s.log"""%( variaveis['banco'], os.path.join( os.path.realpath('.'), 'ctl', arq_ctl ), arq_ctl[:-4] )
        # comando = """sqlldr gfcarga/vivo2019@%s %s > log.log"""%( configuracoes.banco, os.path.join( os.path.realpath('.'), 'ctl', arq_ctl ) )
        
        comando = """sqlldr %s/%s@%s %s > log.log"""%(configuracoes.userBD, configuracoes.pwdBD, configuracoes.banco, arq_ctl )
        os.system(comando)

        ## Buscando os erros ocorridos.
        txt = "-"*100
        txt += """
Carregando arquivo : %s
"""%(arq)
        os.system("cat %s.log | grep -e 'ORA-' -e 'Rejected' -e 'SQL.Loader-' | head -20 > Resumo_%s.log"%(arq_ctl[:-4], arq_ctl[:-4])) 
        fd = open("Resumo_%s.log"%(arq_ctl[:-4]), 'r' )
        erros = fd.readlines()
        fd.close()
        if len(erros) > 0 :
            txt += """
Erros encontrados :
   """
            txt += '   '.join( x for x in erros )
            retorno = False

        ## Gerando resumo da execucao.
        os.system("cat %s.log | grep -e Row -e 'Total logical records' > Resumo_%s.log"%(arq_ctl[:-4], arq_ctl[:-4])) 
        fd = open("Resumo_%s.log"%(arq_ctl[:-4]), 'r' )
        resumo = fd.readlines()
        fd.close()
        txt += """
******************************************************
*** Resumo da execução : ( %s )
*** """%(arq)
        txt += '*** '.join( x for x in resumo )
        txt += '******************************************************\n'

        txt += "-"*100
        txt += "\n"
        log(txt)

        try :
            os.remove( "Resumo_%s.log"%(arq_ctl[:-4]) )
        except :
            pass
    time.sleep(1)
    if not retorno :
        os.system( '> ERRO_%s'%(pid) )
    return retorno   


def encodingDoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r', encoding='iso-8859-1')
        t = fd.read()
        fd.close()
    except :
        return 'utf-8'
    return 'iso-8859-1'


if __name__ == "__main__":
    name_script = os.path.basename(sys.argv[0]).replace('.py', '')
    log("-"*150)
    ret = 0
    txt = ''
    try:
        comum.carregaConfiguracoes(configuracoes)
        comum.addParametro( 'UF', '-UF', "UF dos dados a carregar" , True , 'SP' )
        comum.addParametro( 'VERSAO', '-V', "Nome do diretorio do pleito (Versao)" , True , 'pleito_202011_versao_202011' )

        if not comum.validarParametros() :
            log('### ERRO AO VALIDAR OS PARÂMETROS')
            ret = 91
        else:
            variaveis['UF']         = comum.getParametro('UF')
            variaveis['VERSAO']     = comum.getParametro('VERSAO')
                
            log('UF a processar ..............: %s'%( variaveis['UF'] ))
            log('Pleito ( Versao ) ...........: %s'%( variaveis['VERSAO'] ))

            if not ret :
                try:
                    ret = processar()
                except Exception as e:
                    txt = traceback.format_exc()
                    log(str(e))
                    raise Exception(str(e))
                        
    except Exception as e:
        txt = traceback.format_exc()
        log("ERRO = " + str(e))
        ret = 2
    
    log("="*150)
    sts = 'SUCESSO' if ret == 0 else 'ERRO'
    log( "Status Final ...: " + sts )
    sys.exit(ret)
