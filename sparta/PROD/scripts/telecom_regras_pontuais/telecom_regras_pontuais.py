#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-

"""
- Modulo .......: Teshuvá 
- Jira .........: Teshuvá/PTITES-913 DV - Aba SANEAMENTO / Regras Pontuais Telecom/PTITES-914
- Data Criacao..: 26/10/2021
- Autora........: Fabrisia Gabriela Rosa/ Kyros Tecnonlogia/ fabrisiag@kyros.com.br
- Descricao.....: Este script possibilita a execução dos objetos de banco (procedures) que 
-                 executam regras pontuais (saneador) de telecom, através do painel de execuções. 
----------------------------------------------------------------------------------------------
HISTÓRICO: 
    26/10/2021 : Fabrisia G. Rosa 
        - Criação do Script.

"""

import sys
import os
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)
import configuracoes
import comum
import sql
import datetime
import traceback

comum.carregaConfiguracoes(configuracoes)
log.gerar_log_em_arquivo = True

def fnc_fake_global():
    """
        Função falsa que armazena as variavies globais
    """
    pass 

def fnc_database_connect():
    """
        Função que conecta no banco de dados. 
    """
    try:
        fnc_fake_global.connection = sql.geraCnxBD(configuracoes)

        test_sql = """ SELECT 'PAINEL_EXECUCACAO_' || TO_CHAR(SYSDATE, 'DDMMYYYY_HH24MISS') AS NM_JOB 
                       FROM DUAL       
                   """

        fnc_fake_global.connection.executa(test_sql)

        cursor = fnc_fake_global.connection.fetchone()
        
        if (cursor):
            for index in cursor:
                log(str(index) + " CONEXAO DE SUCESSO COM O BANCO DE DADOS! ")
                fnc_fake_global.nm_job = str(index)
                return 0 
        else:
            log(" ERRO AO CONECTAR COM O BANCO DE DADOS! ")
            return 91 
    except Exception as err:
        err_desc_trace = traceback.format_exc()
        log(" ERRO AO CONECTADO COM O BANCO DE DADOS: " + str(err) + " - TRACE - " + err_desc_trace)
        return 91 

def fnc_valida_parametros():
    """
        Cria e valida os parametros de entrada
    """
    try:
        retorno = 0 
        log("-"*150)

        #Cria
        comum.addParametro('P_TP_TRANSACAO', None, 'Tipo de transacao do banco de dados a ser executado.', True, 'COMMIT')
        comum.addParametro('P_CC_REGRAS', None, 'Nome da regras cadastradas na tabela [GFCADASTRO.TSHTB_CONTROLE_REGRA], usar vírgula para sepação.', True,'SANMC_RN_010, SANMC_RN_030')
        comum.addParametro('P_DT_FILTRO_INICIO', None, 'Data de inicio para realizacao da pesquisa no banco. FORMATO (DD/MM/YYYY).', True,'01/01/2015')        
        comum.addParametro('P_DT_FILTRO_FIM', None, 'Data fim para realizacao da pesquisa no banco. FORMATO (DD/MM/YYYY).', True,'31/01/2015')
        comum.addParametro('P_CC_FILTRO_UF', None, 'Unidade Federativa do Brasil. Aceita apenas um valor.', True,'SP')
        comum.addParametro('P_CC_FILTRO_EMPRESA', None, 'Codigo da Empresa, campo para pesquisa no banco. Aceita mais de um valor separado por virgula.', True,'TBRA')
        comum.addParametro('P_CC_FILTRO_IE', None, 'Inscricao Estadual, campo para pesquisa no banco.Aceita mais de um valor separado por virgula.', False,'108383949112, 999999')
        comum.addParametro('P_CC_FILTRO_FILIAL', None, 'Codigo da Filial, campo para pesquisa no banco. Aceita mais de um valor separado por virgula.', False,'0001, 0002')
        comum.addParametro('P_CC_FILTRO_MODELO', None, 'Codigo do Modelo da NF, campo para pesquisa no banco. Aceita mais de um valor separado por virgula.', False,'21,22')
        comum.addParametro('P_CC_FILTRO_SERIE', None, 'Código da Série da NF, campo para pesquisa no banco. Aceita mais de um valor separado por virgula..', False,'U  T, 06, 1')
        comum.addParametro('P_CC_FILTRO_NOTA', None, 'Número da NF, campo para pesquisa no banco. Aceita mais de um valor separado por virgula.', False,'000000009, 121232222, 999999')
        comum.addParametro('P_CC_FILTRO_OUTROS_FILTROS', None, 'Campo auxiliar, campo para pesquisa no banco Aceita apenas um valor.', False,'ROWNUM > 2 AND ROWNUM < 100')

        #Valida
        if not comum.validarParametros():
            retorno = 91
        else:
            fnc_fake_global.transacao          = comum.getParametro('P_TP_TRANSACAO').upper().strip()
            fnc_fake_global.regras             = comum.getParametro('P_CC_REGRAS').upper().strip()
            fnc_fake_global.filtro_dt_inicio   = comum.getParametro('P_DT_FILTRO_INICIO').upper().strip()
            fnc_fake_global.filtro_dt_fim      = comum.getParametro('P_DT_FILTRO_FIM').upper().strip()
            fnc_fake_global.filtro_uf          = comum.getParametro('P_CC_FILTRO_UF').upper().strip()
            fnc_fake_global.filtro_empresa     = comum.getParametro('P_CC_FILTRO_EMPRESA').upper().strip()
            
            
            try:
                fnc_fake_global.filtro_ie       = comum.getParametro('P_CC_FILTRO_IE').upper().strip()
            except:
                fnc_fake_global.filtro_ie      = ""

            try:    
                fnc_fake_global.filtro_filial   = comum.getParametro('P_CC_FILTRO_FILIAL').upper().strip()
            except:
                fnc_fake_global.filtro_filial  = "" 

            try:
                fnc_fake_global.filtro_modelo   = comum.getParametro('P_CC_FILTRO_MODELO').upper().strip()
            except:
                fnc_fake_global.filtro_modelo  = "" 
            
            try:
                fnc_fake_global.filtro_serie    = comum.getParametro('P_CC_FILTRO_SERIE').upper().strip()
            except:
                fnc_fake_global.filtro_serie   = "" 
            
            try:
                fnc_fake_global.filtro_nota     = comum.getParametro('P_CC_FILTRO_NOTA').upper().strip()
            except:
                fnc_fake_global.filtro_nota    = "" 

            try:
                fnc_fake_global.outros_filtros  = comum.getParametro('P_CC_FILTRO_OUTROS_FILTROS').upper().strip()
            except:
                fnc_fake_global.outros_filtros = "1 = 1" 

            if not retorno:
                try:
                    log("PASSO_1")
                    if (len(fnc_fake_global.filtro_dt_inicio) != 10):
                        log("PASSO_2")
                        log("PARAMETRO DATA_INICIO: Inválido! " + fnc_fake_global.filtro_dt_inicio)
                        retorno = 91
                    else:
                        log("PASSO_3")
                        if (int(fnc_fake_global.filtro_dt_inicio[0:2]) > 31 or int(fnc_fake_global.filtro_dt_inicio[0:2]) < 1): 
                            log("PARAMETRO DATA_INICIO [DIA]: Inválido! " + fnc_fake_global.filtro_dt_inicio[0:2])
                            retorno = 91
                        elif (int(fnc_fake_global.filtro_dt_inicio[3:5]) > 12 or int(fnc_fake_global.filtro_dt_inicio[3:5]) < 1): 
                            log("PARAMETRO DATA_INICIO [MES]: Inválido! " + fnc_fake_global.filtro_dt_inicio[3:5])
                            retorno = 91
                        else:
                            try:
                                log("PARAMETRO DATA_INICIO COM SUCESSO.")
                                fnc_fake_global.data_inicio = datetime.datetime(int(fnc_fake_global.filtro_dt_inicio[6:10]) \
                                                             , int(fnc_fake_global.filtro_dt_inicio[3:5]) \
                                                             , int(fnc_fake_global.filtro_dt_inicio[0:2]))         
                            except:
                                log("PASSO_4")
                                log("PARAMETRO DATA_INICIO: Inválido! " + fnc_fake_global.filtro_dt_inicio)
                                retorno = 91
                except:
                    log("PASSO_5")
                    log("PARAMETRO DATA_INICIO: Inválido! " + fnc_fake_global.filtro_dt_inicio)
                    retorno = 91

            if not retorno:
                try:
                    if (len(fnc_fake_global.filtro_dt_fim) != 10):
                        log("PARAMETRO DATA_FIM: Inválido! " + fnc_fake_global.filtro_dt_fim)
                        retorno = 91
                    else:
                        if (int(fnc_fake_global.filtro_dt_fim[0:2]) > 31 or int(fnc_fake_global.filtro_dt_fim[0:2]) < 1): 
                            log("PARAMETRO DATA_FIM [DIA]: Inválido! " + fnc_fake_global.filtro_dt_fim[0:2])
                            retorno = 91
                        elif (int(fnc_fake_global.filtro_dt_fim[3:5]) > 12 or int(fnc_fake_global.filtro_dt_fim[3:5]) < 1): 
                            log("PARAMETRO DATA_FIM [MES]: Inválido! " + fnc_fake_global.filtro_dt_fim[0:2])
                            retorno = 91
                        else:
                            try:
                                fnc_fake_global.data_fim = datetime.datetime(int(fnc_fake_global.filtro_dt_fim[6:10]) 
                                                             , int(fnc_fake_global.filtro_dt_fim[3:5]) 
                                                             , int(fnc_fake_global.filtro_dt_fim[0:2]))         
                            except:
                                log("PARAMETRO DATA_INICIO: Inválido! " + fnc_fake_global.filtro_dt_fim)
                                retorno = 91
                except:
                    log("PARAMETRO DATA_INICIO: Inválido! " + fnc_fake_global.filtro_dt_fim)
                    retorno = 91    

            if not retorno:
                try:
                    dt_inicio_yyyymmdd = (str(fnc_fake_global.filtro_dt_inicio[6:10])    
                                       + str(fnc_fake_global.filtro_dt_inicio[3:5])  
                                       + str(fnc_fake_global.filtro_dt_inicio[0:2]))

                    dt_fim_yyyymmdd = (str(fnc_fake_global.filtro_dt_fim[6:10])     
                                       + str(fnc_fake_global.filtro_dt_fim[3:5])   
                                       + str(fnc_fake_global.filtro_dt_fim[0:2]))   

                    if int(dt_fim_yyyymmdd) < int(dt_inicio_yyyymmdd):
                        log("PARAMETRO DATA FINAL: Inválido! A data fim não pode ser menor que a data início.")
                        retorno = 91
                except:
                    log("PARAMETRO DATA FINAL: Inválido!" + fnc_fake_global.filtro_dt_fim)   
                    retorno = 91

            if not retorno:
                try:
                    owner_tab_openrisow = getattr(configuracoes, 'owner_openrisow', fnc_fake_global.var_owner_tab_openrisow)
                    if (len(str(owner_tab_openrisow).strip()) > 0):
                        fnc_fake_global.var_owner_tab_openrisow = owner_tab_openrisow
                except:
                    log("PARAMETRO OWNER_OPENRISOW: Inválido! ")
                    retorno = 91 

            if not retorno:
                try:
                    owner_tab_gfcadastro = getattr(configuracoes, 'owner_gfcadastro', fnc_fake_global.var_owner_tab_gfcadastro)
                    if (len(str(owner_tab_gfcadastro).strip()) > 0):
                        fnc_fake_global.var_owner_tab_gfcadastro = owner_tab_gfcadastro
                except:
                    log("PARAMETRO OWNER_GFCADASTRO: Inválido! ")
                    retorno = 91

            if not retorno:
                try:
                    owner_tab_gfcarga = getattr(configuracoes, 'owner_gfcarga', fnc_fake_global.var_owner_tab_gfcarga)
                    if (len(str(owner_tab_gfcarga).strip()) > 0):
                        fnc_fake_global.var_owner_tab_gfcarga = owner_tab_gfcarga
                except:
                    log("PARAMETRO OWNER_GFCARGA: Inválido! ")
                    retorno = 91

        return retorno

    except Exception as err:
        err_desc_trace = traceback.format_exc()
        log("ERRO NA VALIDAÇÃO DOS PARAMETROS DE ENTRADA: " + str(err) + " - TRACE -" + err_desc_trace )
        retorno = 93
        return retorno


def fnc_processar():
    """
        Função que processa as informações
    """
    try:
        retorno = 0 
        log("-"*150)
        log("owner_tab_openrisow  ", fnc_fake_global.var_owner_tab_openrisow)
        log("owner_tab_gfcadastro ", fnc_fake_global.var_owner_tab_gfcadastro)
        log("owner_tab_gfcarga    ", fnc_fake_global.var_owner_tab_gfcarga)

        if fnc_fake_global.connection is not None:
            log("Conexao.............: ", "ATIVA")
            log("Transacao...........: ", fnc_fake_global.transacao)
            log("Job.................: ", fnc_fake_global.nm_job)
            log("Regras..............: ", fnc_fake_global.regras)
            log("Filtro_DT_Inicio....: ", fnc_fake_global.filtro_dt_inicio)
            log("Filtro_DT_Fim.......: ", fnc_fake_global.filtro_dt_fim)
            log("Filtro_UF...........: ", fnc_fake_global.filtro_uf)
            log("Filtro_IE...........: ", fnc_fake_global.filtro_ie)
            log("Filtro_Empresa......: ", fnc_fake_global.filtro_empresa)
            log("Filtro_Filial.......: ", fnc_fake_global.filtro_filial)
            log("Filtro_Modelo.......: ", fnc_fake_global.filtro_modelo)
            log("Filtro_Serie........: ", fnc_fake_global.filtro_serie)
            log("Filtro_Nota.........: ", fnc_fake_global.filtro_nota)
            log("Outros_Filtros......: ", fnc_fake_global.outros_filtros)
            log("-"*150)               

            var_cod_erro = fnc_fake_global.connection.var(int)
            var_desc_erro = fnc_fake_global.connection.var(str)
            var_data_inicio = fnc_fake_global.data_inicio
            var_data_fim = fnc_fake_global.data_fim

            fnc_fake_global.connection.executaProcedure("GFCADASTRO.TSH_SANTL_RP_APLICA_REGRAS",
                                                 P_CD_UF = fnc_fake_global.origem_uf,
                                                 P_NM_JOB = fnc_fake_global.nm_job,
                                                 P_CD_ERRO = var_cod_erro,
                                                 P_DS_ERRO = var_desc_erro,        
                                                 P_DT_FILTRO_INICIO = var_data_inicio,
                                                 P_DT_FILTRO_FIM = var_data_fim,
                                                 P_CC_REGRAS = fnc_fake_global.regras,
                                                 P_CC_FILTRO_UF = fnc_fake_global.filtro_uf,
                                                 P_CC_FILTRO_IE = fnc_fake_global.filtro_ie,
                                                 P_CC_FILTRO_EMPRESA = fnc_fake_global.filtro_empresa,
                                                 P_CC_FILTRO_FILIAL = fnc_fake_global.filtro_filial,
                                                 P_CC_FILTRO_MODELO = fnc_fake_global.filtro_modelo,
                                                 P_CC_FILTRO_SERIE = fnc_fake_global.filtro_serie,
                                                 P_CC_FILTRO_NOTA = fnc_fake_global.filtro_nota,
                                                 P_CC_FILTRO_OUTROS_FILTROS = fnc_fake_global.outros_filtros,
                                                 P_TP_TRANSACAO = fnc_fake_global.transacao,
                                                 P_OWNER_TAB_OPENSISOW = fnc_fake_global.var_owner_tab_openrisow,
                                                 P_OWNER_TAB_GFCADASTRO = fnc_fake_global.var_owner_tab_gfcadastro,
                                                 P_OWNER_TAB_GFCARGA = fnc_fake_global.var_owner_tab_gfcarga
                                                 )

            log(var_cod_erro.getvalue())
            log(var_desc_erro.getvalue())

            try:
                retorno = var_cod_erro.getvalue()
            except:
                retorno = 1
            
            fnc_fake_global.connection.commit()

            return retorno
    
    except Exception as err:
        err_desc_trace = traceback.format_exc()
        log("ERRO NO PROCESSAMENTO: " + str(err) + " - TRACE - " + err_desc_trace)
        retorno = 93
        return retorno 

if __name__ == "__main__":
    """
        Define módulo principal 
    """

    v_retorno = 0 
    err_desc_trace = ''

    fnc_fake_global.connection = None 
    fnc_fake_global.transacao  = ""
    fnc_fake_global.nm_job     = ""
    fnc_fake_global.origem_uf  = "FULL"
    fnc_fake_global.regras     = ""
    fnc_fake_global.filtro_dt_inicio = ""
    fnc_fake_global.filtro_dt_fim    = ""
    fnc_fake_global.filtro_uf = ""
    fnc_fake_global.filtro_ie = ""
    fnc_fake_global.filtro_empresa = ""
    fnc_fake_global.filtro_filial  = ""
    fnc_fake_global.filtro_modelo  = ""
    fnc_fake_global.filtro_serie   = ""
    fnc_fake_global.filtro_nota    = ""
    fnc_fake_global.outros_filtros = ""
    fnc_fake_global.var_owner_tab_openrisow  = "OPENRISOW."
    fnc_fake_global.var_owner_tab_gfcadastro = "GFCADASTRO."
    fnc_fake_global.var_owner_tab_gfcarga    = "GFCARGA."

    try:
        log("-"*150)
        log("INICIO DA EXECUCAO DAS REGRAS PONTUAIS DE TELECOM.".center(120,'#'))
        log(" ")

        if not v_retorno:
            v_retorno = fnc_valida_parametros()
        
        log(" ")
        
        if not v_retorno:
            v_retorno = fnc_database_connect()
        
        log(" ")
        
        if not v_retorno:
            v_retorno = fnc_processar()
        
        log(" ")

        if not v_retorno:
            log("SUCESSO NA EXECUÇÃO! ")
        else:
            log("ERRO NA EXECUÇÃO!")
        
        log(" ")

    except Exception as err:
        err_desc_trace = traceback.format_exc()
        log("ERRO: " + str(err) + " - TRACE - " + err_desc_trace)
        v_retorno = 93

    sys.exit(v_retorno if v_retorno >= log.ret else log.ret)

