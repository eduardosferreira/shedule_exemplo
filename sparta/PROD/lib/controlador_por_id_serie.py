#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: GF
  MODULO ...:
  SCRIPT ...: controlador_por_id_serie.py
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

import configuracoes
import sql


def retornaIdSerieExecutar( mes_ano_inicial, mes_ano_final, uf, series = False, series_not_in = False ) :
    con = sql.geraCnxBD(configuracoes)
    select = """
        SELECT l.ID_SERIE_LEVANTAMENTO
        FROM GFCARGA.TSH_SERIE_LEVANTAMENTO l
        WHERE 1=1
            and to_char(l.MES_ANO, 'yyyymm') >= %s -- passar o mês e ano do parametro mes_ano_inicial
            and to_char(l.MES_ANO, 'yyyymm') <= %s -- passar o mês e ano do parametro mes_ano_final
            and l.UF_FILIAL = '%s' -- parametro UF
            <<SERIES>>
            <<SERIES_NOT_IN>>
    """%( mes_ano_inicial, mes_ano_final, uf )

    clausula_series = '' if not series else "and replace(l.SERIE, ' ', '') in ( '%s' )"%( "','".join( x for x in series.replace("'",'').replace('"','').replace(' ','').split(',')) ) ### os itens das series devem ser separados por virgula e entre aspas'
    clausula_series_not_in = '' if not series_not_in else "and replace(l.SERIE, ' ', '') not in ( '%s' )"%( "','".join( x for x in series_not_in.replace("'",'').replace('"','').replace(' ','').split(',')) ) ### os itens das series devem ser separados por virgula e entre aspas'

    select = select.replace('<<SERIES_NOT_IN>>',clausula_series_not_in)
    select = select.replace('<<SERIES>>',clausula_series)

    ##  Executa o Select
    con.executa(select)
    
    v_lista_id_serie_levantamento = con.fetchall()
    
    return v_lista_id_serie_levantamento


def retiraThreadsFinalizadas(lista_threads, status_thread) :

    v_nova_lista_threads = []
    v_qtde_erros = 0
    for th, id_sparta, idx_thread in lista_threads :
        print('STATUS_THREAD', idx_thread, status_thread[idx_thread])
        if not th.is_alive() :
            # print('>>>>', dir(th))
            print(th.ident, status_thread.get(idx_thread, 'NAO EXISTE'))

            if status_thread[idx_thread] not in ( 0, 1 ) :
                v_qtde_erros += 1
                alteraStatusExecucao(id_sparta, 'ERRO', p_mensagem=False if type(status_thread[idx_thread]) != str else status_thread[idx_thread] )
            else :
                alteraStatusExecucao(id_sparta, 'SUCESSO')

        else :
            v_nova_lista_threads.append( [th, id_sparta, idx_thread] )

    return [ v_nova_lista_threads, v_qtde_erros ]


def criaExecucao(id_sparta, id_serie, script, parametros, usuario, data_prevista = 'SYSDATE' ) :

    if id_sparta > 0 :
        con = sql.geraCnxBD(configuracoes)
        select = "SELECT gfcadastro.SPARTA_SQ_CONTROLADOR_PORSERIE.nextval from dual"
        con.executa(select)
        id_execucao = con.fetchall()[0][0]

        insert = """INSERT INTO gfcadastro.SPARTA_CONTROLADOR_POR_SERIE
( id_execucao, id_serie_levantamento, id_sparta, data_prevista, script, parametros, usuario, status, inicio_execucao )
VALUES
( %s, %s, %s, %s, '%s', '%s', '%s', 'EXECUTANDO', SYSDATE )
        """ %( id_execucao, id_serie, id_sparta, data_prevista, script, parametros, usuario )
        
        con.executa(insert)
        con.commit()
    else :
        id_execucao = 0

    return id_execucao


def alteraStatusExecucao( id_execucao, p_status, p_mensagem = False) :
    if id_execucao > 0 :
        con = sql.geraCnxBD(configuracoes)
        msg = ''
        if p_mensagem :
            msg = ", mensagem = '%s'"%( p_mensagem )
        update = """UPDATE gfcadastro.SPARTA_CONTROLADOR_POR_SERIE 
        SET status = '%s', fim_execucao = SYSDATE %s
        WHERE id_execucao = %s
        """%(p_status, msg, id_execucao)
        con.executa(update)
        con.commit()

    return id_execucao



