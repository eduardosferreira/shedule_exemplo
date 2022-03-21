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

log = comum.log
status_final = 0
log.gerar_log_em_arquivo = True


def processar() :
    banco = configuracoes.banco
    configuracoes.erros = 0
    registros = selectRegistros()
    if isinstance(registros, bool ) :
        return 1
    qtde = len(registros)
    
    max_threads = getattr(configuracoes, 'maximo_processos_paralelos', 3)
    threads_ativas = 0 
    
    lst_status = {}
    configuracoes.lst_status = lst_status

    idx = 0
    if registros :
        while registros :
            if threads_ativas < max_threads :
                reg = registros.pop(0)
                idx += 1
                lst_status[idx] = {}
                # print(reg)

                th = threading.Thread(target = geraConv115 , args= [idx, reg] )
                th.start()
                lst_status[idx]['th'] = th
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
        if isinstance(registros, bool ) :
            return 1 
    
    log("="*80)
    log("Foram realizadas %s execuções."%(qtde))
    log(" - %s com erros."%(configuracoes.erros))
    log(" - %s com sucesso."%(qtde - configuracoes.erros))
    log("="*80)

    return configuracoes.erros


def verificaExecucoes() :
    threads_ativas = 0
    lst_status = configuracoes.lst_status
    for idx in lst_status :
        if lst_status[idx]['status'] == 1 :
            if lst_status[idx]['th'].is_alive() :
                threads_ativas += 1
            else :
                time.sleep(1)
                if lst_status[idx]['status'] == 1 :
                    lst_status[idx]['status'] = 2
        else :
            if lst_status[idx]['status'] == 2 and lst_status[idx]['log'] and not lst_status[idx]['th'].is_alive() :
                log('-'*120)
                log('ERRO - Execução ID < %s > finalizada com ERRO .'%(idx))
                configuracoes.erros += 1
            elif lst_status[idx]['status'] == 0 and lst_status[idx]['log'] and not lst_status[idx]['th'].is_alive() :
                log('-'*120)
                log('Execução ID < %s > finalizada com SUCESSO ! '%(idx))
            log(lst_status[idx]['log'])
            lst_status[idx]['log'] = None
            log('-'*120)

    return threads_ativas


def geraConv115( idx, dados, debug = True ) :
    #### Status da execucação 
    ## 0 - Sucesso
    ## 1 - Executando
    ## 2 - ERRO
    lst_status = configuracoes.lst_status
    ### (16000285, 'SP', datetime.datetime(2016, 4, 1, 0, 0), '0001', '1', 'cd /arquivos/TESHUVA/scripts_rpa/unificado/ &&  nohup ./geraConv115.py 16000285 >log/16000285.log 2>log/16000285.err &', 4)
    id_serie, uf, dt_ref, filial, serie, comando, y = dados

    lst_status[idx]['status'] = 1
    status = 0 

    txt_log = """*"""*100 + '\n'

    if debug :
        txt_log += "** ID serie .........: %s \n"%(id_serie)
        txt_log += "** UF ...............: %s \n"%(uf)
        txt_log += "** Filial ...........: %s \n"%(filial)
        txt_log += "** Serie ............: %s \n"%(serie)
        txt_log += "** Data referencia ..: %s \n"%(dt_ref.strftime('%m/%Y'))
        txt_log += "** Thread execucao ..: %s \n"%(idx)
        txt_log += "** Comando ..........: %s \n"%(comando)
        txt_log += "**\n"
    else :
        txt_log += "Convenio 115 ID: %s Serie: %s Filial: %s Período: %s Iniciado: %s \n"%(id_serie, serie, filial, dt_ref.strftime('%m/%Y'), datetime.datetime.now().strftime('%H:%M:%S') )
    
    ### Executa o comando ...
    # time.sleep(10)
    # status = 2 if idx % 2 == 0 else 0
    status = os.system(comando)
    ### Fim da execução do comando 

    # print('>>>>>', status)
    if status > 0 :
        status = 2
        if debug :
            nome_arq_log = comando.split('>')[1][:-2]
            nome_arq_log_erro = comando.split('>')[2]
            for arq_log in [ nome_arq_log, nome_arq_log_erro ] :
                path_log = os.path.join( comando.split(' ')[1], arq_log )
                if os.path.isfile(path_log) :
                    txt_log += "***** Conteudo do log .: %s \n"%(arq_log)
                    fd = open(path_log, 'r')
                    for l in fd.readlines() :
                        txt_log += "*> %s \n"%(l)
                    fd.close()
                else :
                    txt_log += "***** Log da execucao NAO encontrado ... \n"
                    txt_log += "** Path procurado ..: %s \n"%(path_log)
                txt_log += "*"*100 + '\n'
    
    txt_log += "Convenio 115 ID: %s Serie: %s Filial: %s Período: %s Finalizado com %s : %s \n"%(id_serie, serie, filial, dt_ref.strftime('%m/%Y'), 'ERRO' if status > 0 else 'sucesso',datetime.datetime.now().strftime('%H:%M:%S') )
    txt_log += "*"*100 + '\n'
    
    lst_status[idx]['log'] = txt_log
    lst_status[idx]['status'] = status 
    return True


def selectRegistros() :
    banco = configuracoes.banco
    log("Gerando lista de comandos a processar ...")

    obj_sql = sql.geraCnxBD(configuracoes)
    
    resultado = []
    try:
        comando = """
        select  a.id_serie_levantamento, a.uf, a.mes_ano, FILI_COD, REPLACE(a.serie,' ','') serie,
            'cd /portaloptrib/TESHUVA/sparta/PROD/scripts/Convenio115/gera_conv115'||' && '||' ./gera_conv115.py '||a.id_serie_levantamento||' >log/'||a.id_serie_levantamento||'.log'||' 2>log/'||a.id_serie_levantamento||'.err ' LINHA_COMANDO_CONVENIO_ATUAL,
            row_number()over (order by mes_ano asc,serie asc) ordem_exec_sugerida
        from gfcarga.v_tsh_info_serie_levant_v2 A
        where 1=1
            """

        if comum.getParametro('MesAno') :
            mes, ano = [ comum.getParametro('MesAno')[:2], comum.getParametro('MesAno')[2:] ]

            comando += """and mes_ano >= to_date('01/%s/%s','dd/mm/yyyy')  --parametro: <MESANO>: 
            and mes_ano <  add_months(to_date('01/%s/%s','dd/mm/yyyy'),1)  --parametro:  <MESANO>: 
            """%( mes.rjust(2,'0'), ano, mes.rjust(2,'0') , ano )
            
            if comum.getParametro('Filial') :
                filiais = ", ".join( "'%s'"%(x.strip()) for x in comum.getParametro('Filial').split(',') )
                comando += """AND FILI_COD IN ( %s )  --parametro: <FILI_COD>
            """%( filiais )
                
        else :
            mesI, anoI = [ comum.getParametro('MesAnoInicio')[:2], comum.getParametro('MesAnoInicio')[2:] ]
            mesF, anoF = [ comum.getParametro('MesAnoFim')[:2], comum.getParametro('MesAnoFim')[2:] ]

            comando += """and mes_ano >= to_date('01/%s/%s','dd/mm/yyyy')  --parametro: <MESANOINI>: 
            """%( mesI, anoI )
            comando += """and mes_ano <  add_months(to_date('01/%s/%s','dd/mm/yyyy'),1)  --parametro:  <MESANOFIM>: 
            """%( mesF, anoF )

            if comum.getParametro('Filial') :
                filiais = ", ".join( "'%s'"%(x.strip()) for x in comum.getParametro('Filial').split(',') )
            else :
                filiais = "''"
                
            comando += """AND (FILI_COD IN (%s)  --parametro: <FILI_COD>
            or FILI_COD in (select fili_cod from openrisow.filial where emps_cod = 'TBRA' and fili_cod_insest =  '%s'))  --parametro: <IE>
            """%( filiais, '' if not comum.getParametro('IE') else comum.getParametro('IE') )

        if comum.getParametro('Serie') :
            series = ", ".join( "'%s'"%(x.replace(' ','')) for x in comum.getParametro('Serie').split(',') )
            comando += """and REPLACE(a.serie,' ','') in ( %s )  --parametro: <SERIE>
            """%( series )
            
        comando +="""and REPLACE(a.serie,' ','') not in ('AS1','ASS','AS2','AS3','T1')  --Não é um parâmetro e sim restrição para series que não são processadas
            order by ordem_exec_sugerida
                """

        log("Executando comando : ", comando )

        obj_sql.executa(comando)
        resultado = obj_sql.fetchall() 
        log('- Encontrado(s) %s registro(s) para executar .'%(len(resultado)))
    except Exception as e :
        log('Erro na conexao ao buscar registros. \n - Erro encontrado :', e)
        return False

    return resultado


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
    comum.addParametro('MesAno', '-M', 'Mes e ano de referencia, mês com dois dígitos, ano com quatro dígitos.', False, '012015')
    comum.addParametro('Filial', '-F', 'Filial(is) a serem processadas.', False, '0001, 9144')
    comum.addParametro('Serie',  '-S', 'Serie(s) a serem processadas.', False, 'U K , 1, C')

    ### Parametros utilizados pelo processamento Massivo
    comum.addParametro('MesAnoInicio', '-MI', 'Mes e ano de inicio da referencia, mês com dois dígitos, ano com quatro dígitos.', False, '012015')
    comum.addParametro('MesAnoFim',    '-MF', 'Mes e ano de fim da referencia, mês com dois dígitos, ano com quatro dígitos.', False, '012015')
    comum.addParametro('IE',           '-IE', """Inscricao estadual a ser processada.
    * Obs.: 
      Para executar o processo Massivo por mes, deve-se utilizar os parametros 
        MesAno          -M
        Filial          -F
        Serie           -S
      
      Para executar o processo Massivo para varios meses, deve-se utilizar os parametros 
        MesAnoInicio    -MI
        MesAnoFim       -MF
        IE              -IE
        Filial          -F
        Serie           -S
    
      conforme a descrição de cada um acima.
    """, False, '108383949112')

    if not comum.validarParametros() :
        ret = 3
    
    if not ret and not comum.getParametro('MesAno') :
        if ( not comum.getParametro('MesAnoInicio') ) or ( not comum.getParametro('MesAnoFim') ) :
            log('ERRO - Deve-se informar o mes e o ano de inicio e fim a ser processado.')
            ret = 4
        if not ( comum.getParametro('Filial') or comum.getParametro('IE') ) :
            log('ERRO - Deve-se informar a Filial ou a IE que será processada.')
            ret = 4 
        
    if ret == 4 :
        log( "Erro, parâmetros inválidos !" )
        comum.imprimeHelp()

    return True if not ret else False


if __name__ == "__main__":
    ret = 0 
    if inicializar() :
        if processar() != 0:
            log('ERRO no processamento ... Verifique')
            ret = 1
    else :
        ret = 2
    
    sys.exit(ret)
