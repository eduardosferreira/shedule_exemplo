#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SCRIPT ......: gera_conv115_massivo.py
  CRIACAO .....: 01/02/2021
  AUTOR .......: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO ...: 
  DOCUMENTACAO : 01 - Teshuva_RMSV0_Execução Convenio 115 Por Serie_OLDConsulta.docx
                 01 - Teshuva_RMSV0_Execução Massiva Convenio 115.docx

----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 01/02/2021 - Welber Pena de Sousa - Kyros Tecnologia
        - Criacao do script.

    * 26/08/2021 - Welber Pena de Sousa - Kyros Tecnologia
        - Alteração do script para nova versao e padroes do Painel de execucoes.
----------------------------------------------------------------------------------------------
"""
import sys
import os

SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)

import datetime
import threading
import time

import configuracoes
import comum
import layout
import sql
import agente_sparta

log = comum.log
status_final = 0
log.gerar_log_em_arquivo = True


def retornaIdServidor():
    ### Gera configuracao ... dados servidor webservice
    ips = agente_sparta.retornaIpsLocal()
    ip = ips['bond0']
    port = 65502
    # log('IP:...............', ip)
    # log('PORTA:............', port)
    dados_server = { 'HOST' : ip, 'PORT': port }

    ### Chama funcao
    msg = agente_sparta.execFncXmlRpc( dados_server, 'buscaIdServidor' )
    return msg['id_servidor']


def retornaIdScriptServidor(id_servidor, script):
    ### Gera configuracao ... dados servidor webservice
    ips = agente_sparta.retornaIpsLocal()
    ip = ips['bond0']
    port = 65502
    # log('IP:...............', ip)
    # log('PORTA:............', port)
    dados_server = { 'HOST' : ip, 'PORT': port }

    ### Chama funcao
    msg = agente_sparta.execFncXmlRpc( dados_server, 'buscaIdScriptServidor', script, id_servidor )
    return msg['id_script_servidor'], msg['id_script']


def retornaListaParametros(id_script):
    ### Gera configuracao ... dados servidor webservice
    ips = agente_sparta.retornaIpsLocal()
    ip = ips['bond0']
    port = 65502
    # log('IP:...............', ip)
    # log('PORTA:............', port)
    dados_server = { 'HOST' : ip, 'PORT': port }

    ### Chama funcao
    msg = agente_sparta.execFncXmlRpc( dados_server, 'buscaListaParametros', id_script )
    return msg['parametros']

def buscaExecucoes(id_agenda):
    ### Gera configuracao ... dados servidor webservice
    ips = agente_sparta.retornaIpsLocal()
    ip = ips['bond0']
    port = 65502
    # log('IP:...............', ip)
    # log('PORTA:............', port)
    dados_server = { 'HOST' : ip, 'PORT': port }

    ### Chama funcao
    msg = agente_sparta.execFncXmlRpc( dados_server, 'buscaListaExecucoes', id_agenda )
    return msg['execucoes']


def cadastraExecucao(id_script_servidor, id_user, solicitante, dic_parametros):
    ### Gera configuracao ... dados servidor webservice
    ips = agente_sparta.retornaIpsLocal()
    ip = ips['bond0']
    port = 65502
    # log('IP:...............', ip)
    # log('PORTA:............', port)
    dados_server = { 'HOST' : ip, 'PORT': port }

    ### Chama funcao registraAgenda(self, id_script_servidor, id_user, solicitante, dic_parametros ) 
    msg = agente_sparta.execFncXmlRpc( dados_server, 'registraAgenda', id_script_servidor, id_user, solicitante, dic_parametros )
    print('MSG >>>>', msg)
    return msg


def processar() :
    banco = configuracoes.banco
    configuracoes.erros = 0
    #### OQ varia na execucao individual ????
    ### Neste script o Mes de referencia que deve ser incrementado apartir dos parametros :
    ### MesAnoInicio, '-MI'
    ### MesAnoFim,    '-MF'

    meses = []

    mes_inicial = int(comum.getParametro('MesAnoInicio')[:2])
    ano_inicial = int(comum.getParametro('MesAnoInicio')[2:])
    mes_final   = int(comum.getParametro('MesAnoFim')[:2])
    ano_final   = int(comum.getParametro('MesAnoFim')[2:])
    
    # print(mes_inicial, ano_inicial)
    # print(mes_final, ano_final)

    mes = mes_inicial
    ano = ano_inicial

    while (mes != mes_final) or (ano != ano_final) :
        meses.append( "%s/%s"%( str(mes).rjust(2, '0'), ano ) )
        mes += 1
        if mes > 12 :
            mes = 1
            ano += 1
    
    meses.append( "%s/%s"%( str(mes).rjust(2, '0'), ano ) )

    # print ('MESES', meses )

    max_threads = getattr(configuracoes, 'maximo_processos_paralelos', 1)
    threads_ativas = 0 
    
    lst_status = {}
    configuracoes.lst_status = lst_status

    idx = 0
    qtde = 0
    if meses :
        while meses :
            if threads_ativas < max_threads :
                reg = meses.pop(0)
                idx += 1
                lst_status[idx] = {}
                # print(reg)

                th = threading.Thread(target = criaNovaExecucao , args= [idx, reg] )
                th.start()
                qtde += 1
                lst_status[idx]['th'] = th
                time.sleep(1)
                # criaNovaExecucao(idx, reg)
            else :
                time.sleep(5)
                print('.', end='')
            threads_ativas = verificaExecucoes()
            time.sleep(1)
        ### Aguarda a finalização de todas as instancias 
        log('Aguardando finalização das instancias ...')
        threads_ativas = verificaExecucoes()
        while threads_ativas :
            time.sleep(10)
            threads_ativas = verificaExecucoes()
        log('Todas as execuções finalizadas ...')

    else :
        if isinstance(meses, bool ) :
            return 1 
    
    log("="*80)
    log("Foram realizadas %s execuções."%(qtde))
    log(" - %s com erros."%(configuracoes.erros))
    log(" - %s com sucesso."%(qtde - configuracoes.erros))
    log("="*80)

    return configuracoes.erros


def criaNovaExecucao(idx, reg) :
    # http://10.238.10.208:65500/Sparta/interfaceExecutor/retornaIdServidor?ip=10.238.10.208&h=brtlvlts1198pl
    id_server = retornaIdServidor()
    configuracoes.lst_status[idx]['status'] = 1 ### Executando
    log("[ THREAD", idx,"] - ID SERVIDOR ... :", id_server)
    id_script_servidor, id_script = retornaIdScriptServidor(id_server, 'Loader Conv115 Original Protocolado')
    log("[ THREAD", idx,"] - ID SCRIPT ............... :", id_script)
    log("[ THREAD", idx,"] - ID SCRIPT SERVIDOR ...... :", id_script_servidor)
    status = 0

    lst_parametros = retornaListaParametros(id_script)
    dic_parametros = {}
    parametros_valores = {}
    parametros_valores['UF'] = comum.getParametro('UF')
    parametros_valores['Periodo de dados ( MM/YYYY )'] = reg
    parametros_valores['Serie'] = comum.getParametro('Serie')
    parametros_valores['Processar'] = comum.getParametro('Processar')
    
    erro = 0
    for id_param in lst_parametros :
        valor = ''
        if lst_parametros[id_param] in parametros_valores.keys() :
            valor = parametros_valores[lst_parametros[id_param]]
        elif lst_parametros[id_param] == 'ID Série' :
            # valor = reg[0]
            valor = 1
        else :
            erro += 1
            log("[ THREAD", idx,"] - ERRO - Falta valor para o parametro :", id_param, lst_parametros[id_param])

        dic_parametros[id_param] = valor
        log("[ THREAD", idx,"] -",id_param.ljust(10,'.'), ':', lst_parametros[id_param], '=', valor)
    
    if erro > 0 :
        return False

    resultado = cadastraExecucao( id_script_servidor, 2, os.path.basename(sys.argv[0]), dic_parametros )
    if resultado['status'] != 'Ok' :
        log('ERRO ao cadastrar execucao ... ')
        log(resultado)
        status = 1
    
    if not status :
        # id_agenda = 4912
        id_agenda = resultado.get('id_agenda', False)
        log("[ THREAD", idx,"] - ID AGENDA ............... :", id_agenda)
        log('-'*100)
        if id_agenda :
            #### Aguarda execucao da agenda cadastrada ....
            lst_execucoes = buscaExecucoes(id_agenda)
            log("[ THREAD", idx,"] - >>>>", lst_execucoes)

        #     dic_status_execucao = {
        #     0 : 'Cancelada',
        #     1 : 'Pendente',
        #     2 : 'Fila de execução',
        #     3 : 'Atrasada',
        #     4 : 'Executando',
        #     5 : 'Sucesso',
        #     6 : 'Erro',
        #     7 : 'Morto',
        #     8 : 'Encadeamento' }

            id_execucao = 0
            aguardar = True
            while aguardar :
                for k in lst_execucoes :
                    if int(k) > id_execucao :
                        id_execucao = int(k)
                        aguardar = False
                time.sleep(10)
                lst_execucoes = buscaExecucoes(id_agenda)

            aguardar = True
            log("[ THREAD", idx,"] - Aguardando ... processo", id_execucao)
            while aguardar :

                if lst_execucoes[str(id_execucao)][1] in [ 'Sucesso', 'Erro', 'Morto' ] :
                    log("[ THREAD", idx,"] - Execucao : %s finalizada com status %s"%(id_execucao, lst_execucoes[str(id_execucao)][1]))
                    aguardar = False
                    status = 2
                else :
                    time.sleep(5)
                    
                    lst_execucoes = buscaExecucoes(id_agenda)
            
            
            if lst_execucoes[str(id_execucao)][0] != 5 :
                status = lst_execucoes[str(id_execucao)][0]
        else :
            log("[ THREAD", idx,"] - ERRO ao cadastrar execucao ...")
            # log("[ THREAD", idx,"] -", resultado)
            status = 2

    time.sleep(15)    
    configuracoes.lst_status[idx]['status'] = status ### Finalizado com SUCESSO
    return True


def verificaExecucoes() :
    threads_ativas = 0
    lst_status = configuracoes.lst_status
    for idx in lst_status :
        if lst_status[idx]['status'] == 1 :
            if lst_status[idx]['th'].is_alive() :
                threads_ativas += 1
            else :
                time.sleep(2)
                if lst_status[idx]['status'] == 1 :
                    lst_status[idx]['status'] = 2
        elif lst_status[idx]['status'] < 3 :
            if lst_status[idx]['status'] == 2 and not lst_status[idx]['th'].is_alive() :
                log('-'*120)
                log('ERRO - Execução ID < %s > finalizada com ERRO .'%(idx))
                configuracoes.erros += 1
                lst_status[idx]['status'] = 3
            elif lst_status[idx]['status'] == 0 and not lst_status[idx]['th'].is_alive() :
                log('-'*120)
                log('Execução ID < %s > finalizada com SUCESSO ! '%(idx))
                lst_status[idx]['status'] = 3
            log('-'*120)

    return threads_ativas


def inicializar() :
    comum.carregaConfiguracoes(configuracoes)
    ret = 0
    configuracoes.dic_layouts = layout.carregaLayout()
    if not configuracoes.dic_layouts :
        ret = 2
    
    if not getattr(configuracoes, 'banco', False) :
        log("Erro falta variavel 'banco' no arquivo de configuração (.cfg).")
        ret = 1

    #### Reconhecer parametros passados :
    # <MESANO>: Obrigatório, mês com dois dígitos, ano com quatro dígitos.
    # <FILI_COD>: Opcional, exemplo: 0001,9144
    # <SERIE>: Opcional, exemplo: UK,1,C 
    #
    #   OBS.: Os parâmetros <SERIE> e <FILI_COD> aceitam listas como parâmetro, 
    #   recomendo que seja alinhado com o time um padrão para a passagem dos parâmetros. 
    #   Sugiro que seja implementado que as listas sejam separadas por vírgula e sem espaço 
    #   (retirando inclusive os espaços das series).
    comum.addParametro('UF', '-UF', 'UF a ser processada.', False, 'SP')
    comum.addParametro('Serie',  '-S', 'Serie(s) a serem processadas.', False, 'U K , 1, C')

    ### Parametros utilizados pelo processamento Massivo
    comum.addParametro('MesAnoInicio', '-MI', 'Mes e ano de inicio da referencia, mês com dois dígitos, ano com quatro dígitos.', True, '012015')
    comum.addParametro('MesAnoFim',    '-MF', 'Mes e ano de fim da referencia, mês com dois dígitos, ano com quatro dígitos.', False, '012015')
    comum.addParametro('Processar',    '-P', """Utilizado para dizer o que sera carregado : ( Opcoes : Mestre, Item, Cadastro, Todos )""", False, 'Todos')

    if not comum.validarParametros() :
        ret = 3
    
    # if ( not ret and not comum.getParametro('MesAno') ) and not comum.getParametro('IdSerie') :
    #     if ( not comum.getParametro('MesAnoInicio') ) or ( not comum.getParametro('MesAnoFim') ) :
    #         log('ERRO - Deve-se informar o mes e o ano de inicio e fim a ser processado.')
    #         ret = 4
    #     if not ( comum.getParametro('Filial') or comum.getParametro('IE') ) :
    #         log('ERRO - Deve-se informar a Filial ou a IE que será processada.')
    #         ret = 4 
        
    if ret == 4 :
        log( "Erro, parâmetros inválidos !" )
        comum.imprimeHelp()
    
    return True if not ret else False


if __name__ == "__main__":
    ret = 0 
    if inicializar() :
        if not processar() :
            log('ERRO no processamento ... Verifique')
            ret = 1
    else :
        ret = 2
    
    sys.exit(ret)
