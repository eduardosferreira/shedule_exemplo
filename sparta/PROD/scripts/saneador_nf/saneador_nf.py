#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-

"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: aplica_regras.py
  CRIACAO ..: 22/06/2021
  AUTOR ....: Victor Santos Cardoso / KYROS Consultoria
  DESCRICAO : 
----------------------------------------------------------------------------------------------
  HISTORICO :
    * 01/06/2021 - Victor Santos Cardoso / KYROS Consultoria - Criacao do script.

----------------------------------------------------------------------------------------------
"""

import sys
import os
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)
import configuracoes
import datetime

import cx_Oracle

import comum
import sql

def APLICA_REGRAS(job,uf,dt_ini,dt_fim,filtro,hint,regra):
 
    connection = sql.geraCnxBD(configuracoes)
    cd_erro = connection.var(int)
    ds_erro = connection.var(str)
    fl_opcao = 2
    commit = 'COMMIT'
    p_idcontrole_procs= connection.var(int)
    # procedure  = "SPT80732427.TSH_SANTL_40220_APLICA_REGRAS" 
    procedure  = "gfcadastro.TSH_SANTL_40220_APLICA_REGRAS" 
    parametros = [  
                    job,
                    uf,
                    dt_ini,  ### p_dt_ini    -- periodo inicial da tabela mestre telcom (notas)
                    dt_fim,  ### p_dt_fim    -- periodo final da tabela mestre telcom (notas)
                    p_idcontrole_procs,
                    cd_erro,
                    ds_erro,
                    fl_opcao,
                    filtro,  ### p_cc_filtro -- filtros dos campos da tabela de nota (com alias nf. Ex. nf.emps_cod = 'TBRA' and nf.emps_cod = 'TBRA')
                    hint,    ### p_cc_hint   -- caso deseja que colocar algum hint de processamento, como por exemplo parallel(8) -- varificar
                    regra,   ### p_cc_regra  -- nomes das regras do saneados, vinculados a tabela "TSHTB_CONTROLE_REGRA". ### Exemplo : SANTL_RN_010, SANTL_RN_020, ambos pode ser repassados separados por virgula
                    commit   ### p_cc_commit -- forma de transação, commit ou rollback
                ]
    

    # log('# Executando procedure ..: %s'% procedure)
    # log('# Com parametros ....: ' + ', '.join( str(x) for x in parametros ))
    connection.executaProcedure(procedure, *parametros)
    return [cd_erro.getvalue(), ds_erro.getvalue(), p_idcontrole_procs.getvalue()]

def dtf():
    return (datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'))

def dtfH():
    return (datetime.datetime.now().strftime('%d/%m/%Y'))    

def ultimodia(ano,mes):
    return(31 if mes == 12 else datetime.date.fromordinal((datetime.date(ano,mes+1,1)).toordinal()-1).day)    

def sumarizacao(idcontrole):
    global print_query
    query="""SELECT
                SUM(a.qt_entrada)               qtd_nf,
                SUM(a.qt_nf_atualizadas)        qtd_nf_atua,
                SUM(a.qt_inf_atualizados)       qtd_inf_atua
              FROM
                gfcadastro.tshtb_controle_jobs a where CD_EXEC_EXTERNAL = '%s'
    """%(idcontrole)

    print_query = query
    
    cursor = sql.geraCnxBD(configuracoes)
    cursor.executa(query)
    result = cursor.fetchone()

    if result == None:
        print("#### ATENÇÃO: NENHUM RESULTADO PARA SUMARIZAÇÃO DOS DADOS")
        print("####     Query = ")
        print("####")
        print(query)
        print("####")
        ret=99
        return(ret)

    return(result)

def processar():
    global print_query
    log('# DADOS INPUT - ',sys.argv)
    log("# ")
    v_job    = ""
    v_uf     = ""
    v_mes    = ""
    v_ano    = ""
    v_dt_ini = ""
    v_dt_fim = ""
    cd_erro  = ""
    ds_erro  = ""
    v_filtro = ""
    v_hint   = "noparallel"
    v_regra  = None

    if (len(sys.argv) > 2 ): 
        v_job    = sys.argv[0].upper()
        v_uf     = sys.argv[1].upper()
        v_mes    = sys.argv[2][:2].upper()
        v_ano    = sys.argv[2][2:].upper()
        v_dt_ini = datetime.datetime(int(v_ano), int(v_mes), 1)
        ult_dia  = datetime.datetime(int(v_ano),int(v_mes),int(ultimodia(int(v_ano),int(v_mes))))
        v_dt_fim =  ult_dia
        if (len(sys.argv) > 3 ):
            if sys.argv[3] == "":
                v_filtro = "1=1"
            else:
                v_filtro = sys.argv[3].upper()
        if (len(sys.argv) > 4 ):
            v_regra  = sys.argv[4].upper()
    else:
        log("-" * 100)
        log("#### ")
        log('#### ERRO - Erro nos parametros do script.')
        log("#### ")
        log('#### Exemplo de como deve ser :')
        log('####      %s <UF> <MMAAAA> [FILTRO] [REGRA]... '%(sys.argv[0]))
        log("#### ")
        log('#### Onde')
        log('####      <UF> = estado. Ex: SP')
        log('####      <MMAAAA> = mês e ano. Ex: Para Junho de 2020 informe 062020')
        log('####      [FILTRO] = É opcional, pode ou não ser informado.')
        log("####                 EXEMPLO: nf.emps_cod = 'TBRA' and nf.FILI_cod = '0001' - É necessário informar NF. " )
        log('####      [REGRA] = É opcional, pode ou não ser informado.')
        log('####                Nome das regras do saneador, vinculados a tabela "TSHTB_CONTROLE_REGRA".') 
        log('####                Exemplo : SANTL_RN_010, SANTL_RN_020, se deseja mais de 1 regra, informe separando por virgula.')
        log("#### ")
        log('#### Portanto, se o estado = SP, o mes = 06 e o ano = 2020, e deseja incluir filtro e regra, o comando correto deve ser :')  
        log('####      %s SP 012015 nf.emps_cod = "TBRA" "SANTL_RN_010, SANTL_RN_020" '%(sys.argv[0]))  
        log("#### ")
        log('#### Se desejar processar todo o mês, o comando deve ser:')  
        log('####      %s SP 062020'%(sys.argv[0]))  
        log("#### ")
        log("-" * 100)
        log("")
        log("Retorno = 99") 
        ret = 99
        return(99)

    log("# ",dtf() , "PROCESSO SENDO STARTADO...")
    log("# ")
    cd_erro,ds_erro,idcontrole_procs = APLICA_REGRAS(v_job,v_uf,v_dt_ini,v_dt_fim,v_filtro,v_hint,v_regra)
    log ('CODIGO      -> ', cd_erro)
    log ('DESCRIÇÃO   -> ', ds_erro)
    log ('ID CONTROLE -> ', idcontrole_procs)
    log("-")
    if cd_erro != None and cd_erro > 0:
        log ('### ERRO NA EXECUÇÃO DO SCRIPT')
        log ("###", ds_erro)
        return(99)
    else:
        log("# ",dtf() , "EXECUÇÃO FINALIZADA COM SUCESSO...") 
        log("-" * 100)
        total = sumarizacao(idcontrole_procs)
        # print(total)
        log ('QUANTIDADE DE ENTRADA           -> ', total[0])
        log ('QUANTIDADE DE NOTAS ATUALIZADAS -> ', total[1])
        log ('QUANTIDADE DE ITENS ATUALIZADOS -> ', total[2])
        log("-" * 100)
        log("CONFIRA TAMBÉM O RESULTADO EM: ")
        log(print_query)
        return(0)

if __name__ == "__main__":
    cod_saida = 0
    log('-'*100)
    log("# ")  
    log("# ",dtf() , " - Início do processamento.")
    log("# ")
    comum.carregaConfiguracoes(configuracoes)
    cod_saida = processar()
    log('-'*100)
    log("Código de saida = ",cod_saida)
    sys.exit(cod_saida)