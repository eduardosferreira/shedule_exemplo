#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
MODULO ...: TESHUVA
SCRIPT ...: saneador_notas_mercadoria.py
CRIACAO ..: 28/07/2021
AUTOR ....: EDUARDO DA SILVA FERREIRA / KYROS TECNOLOGIA
            eduardof@kyros.com.br
DESCRICAO.: Mercadoria - Saneador Notas Fiscais
----------------------------------------------------------------------------------------------
PARAMETROS: 
Parâmetros de entrada:
1 ) P_TP_TRANSACAO               : [OBRIGATÓRIO  ] Tipo de transacao do banco de dados a ser executado . Exemplo: COMMIT
2 ) P_CC_ORIGEM_UF               : [OBRIGATÓRIO  ] Referente se sera acessado as tabelas oficiais ou saneadas . Exemplo: FULL
3 ) P_CC_REGRAS                  : [OBRIGATÓRIO  ] Nome das regras cadastradas na tabela [GFCADASTRO.TSHTB_CONTROLE_REGRA] separados por virgula. Exemplo: SANMC_RN_010, SANMC_RN_030
4 ) P_DT_FILTRO_INICIO           : [OBRIGATÓRIO - FORMATO DD/MM/YYYY] Data de inicio para realizacao da pesquisa . Exemplo: 01/01/2015
5 ) P_DT_FILTRO_FIM              : [OBRIGATÓRIO - FORMATO DD/MM/YYYY] Data de fim    para realizacao da pesquisa . Exemplo: 31/01/2015
6 ) P_CC_FILTRO_UF               : [OBRIGATÓRIO ] Unidade Federativa do Brasil separados por virgula . Exemplo: SP, RJ, PR
7 ) P_CC_FILTRO_EMPRESA          : [OBRIGATÓRIO ] Codigo da Empresa, campo para pesquisa separado por virgula. Exemplo. TBRA, 0001
8 ) P_CC_FILTRO_IE               : [OPCIONAL    ] Inscricao Estadual, campo para pesquisa separado por virgula. Exemplo. 108383949112, 999999
9 ) P_CC_FILTRO_FILIAL           : [OPCIONAL    ] Codigo da Filial, campo para pesquisa separado por virgula. Exemplo. 0001, 0002
10) P_CC_FILTRO_MODELO           : [OPCIONAL    ] Codigo do Modelo da NF, campo para pesquisa separado por virgula. Exemplo. 21,22
11) P_CC_FILTRO_SERIE            : [OPCIONAL    ] Código da Serie da NF, campo para pesquisa separado por virgula. Exemplo. U  T, 06, 1
12) P_CC_FILTRO_NOTA             : [OPCIONAL    ] Numero da NF, campo para pesquisa separado por virgula. Exemplo. 000000009, 121232222, 999999
13) P_CC_FILTRO_OUTROS_FILTROS   : [OPCIONAL    ] Campo auxiliar, campo para pesquisa nao separado por virgula. Exemplo. ROWNUM > 2 AND ROWNUM < 100

----------------------------------------------------------------------------------------------
    HISTORICO : 
        * 28/07/2021 - EDUARDO DA SILVA FERREIRA / KYROS TECNOLOGIA (eduardof@kyros.com.br)
        - Criacao do script.
        * 30/08/2021 - EDUARDO DA SILVA FERREIRA / KYROS TECNOLOGIA (eduardof@kyros.com.br)
        - PTITES-131 : DV - Novo Padrão: SANEAMENTO / Mercadoria - Mercadoria - Regras Pontuais
        - Alterado nome para : saneador_notas_mercadoria.py
        
        * 28/09/2021 - AIRTON BORGES / KYROS TECNOLOGIA (airtonb@kyros.com.br)
        - Copiar o script saneador_notas_mercadoria.py para um novo script com o nome aplica_ajustes_manuais.py 
        - Alterar os parâmetros de entrada do novo script aplica_ajustes_manuais.py :
            1.	Transação : Retirar esse parâmetro do script, mas internamente na estrutura do script o seu valor continua mas com o valor fixo de "COMMIT".
            2.	Regras : Manter esse parâmetro sendo obrigatório mas o mesmo não deve ter valor padrão, para forçar o usuário a especificar um dos valores da tabela abaixo.
            3.	Data_inicio : Retirar esse parâmetro do script, mas internamente na estrutura do script o seu valor continua mas com o valor fixo de "01/01/1990".
            4.	Data_fim : Retirar esse parâmetro do script, mas internamente na estrutura do script o seu valor continua mas com o valor fixo de "01/01/2050".
            5.	UF : Retirar esse parâmetro do script.
            6.	Empresa : Retirar esse parâmetro do script, mas internamente na estrutura do script o seu valor continua mas com o valor fixo de "TBRA".
            7.	IE : Retirar esse parâmetro do script.
            8.	Filial : Retirar esse parâmetro do script.
            9.	Modelo : Retirar esse parâmetro do script.
            10.	Serie : Retirar esse parâmetro do script.
            11.	Nota : Retirar esse parâmetro do script.
            12.	Outros : Retirar esse parâmetro do script, mas internamente na estrutura do script o seu valor continua mas com o valor fixo de "TMP.GAP = ‘<Nome do GAP>’"
            13.	GAP : Esse parâmetro deve ser incluído no script e o valor recebido nele deve ser passado para o parametro anterior outros no lugar de <Nome do GAP> .

        
----------------------------------------------------------------------------------------------
"""

import sys
import os
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)
import configuracoes
from comum import log
import comum
import sql
import datetime
import traceback

# Demais configuração
comum.carregaConfiguracoes(configuracoes)
log.gerar_log_em_arquivo = True


def ob_global():
    """
        Funcao falsa apenas para sergir de apoio para armazenar as variaveis globais 
    """
    pass

def fnc_conectar_banco_dados():
    """
        Funcao para conectar na base de dados 
    """
    try:
        
        ob_global.gv_ob_conexao = sql.geraCnxBD(configuracoes)
        v_ds_sql="""
        SELECT 'PAINELEXECUCAO_'||TO_CHAR(SYSDATE,'YYYYMMDD_HH24MISS') NM_JOB FROM DUAL
        """
        ob_global.gv_ob_conexao.executa(v_ds_sql)
        
        v_ob_cursor = ob_global.gv_ob_conexao.fetchone()
        if (v_ob_cursor):    
            for campo in v_ob_cursor:
                log(str(campo) + " SUCESSO CONEXAO BANCO DE DADOS") 
                ob_global.gv_nm_job = str(campo)
                return 0
                break
        
        else:
            log("ERRO CONEXAO BANCO DE DADOS ") 
            return 91
        
    except Exception as e:
        v_ds_trace = traceback.format_exc()
        log("ERRO CONEXAO BANCO DE DADOS .: " + str(e) + " - TRACE - " + v_ds_trace)
        return 91


def fnc_validar_entrada():
    """
        Retorna a validação de entrada e dos arquivos de configuração
    """
    try:
        v_nr_retorno = 0        
        log("-"*150)

        comum.addParametro('P_CC_REGRAS', None, 'Nome das regras cadastradas na tabela [GFCADASTRO.TSHTB_CONTROLE_REGRA] separados por virgula.', True,'')
        comum.addParametro('P_CC_GAP', None, 'Nome do GAP', True,'')
       
        # Validacao dos parametros de entrada
        if not comum.validarParametros() :
            v_nr_retorno = 91
        else:
            ob_global.gv_tp_transacao              = "COMMIT"   
            #ob_global.gv_tp_transacao              = "ROLLBACK"   
            ob_global.gv_cc_regras                 = comum.getParametro('P_CC_REGRAS').upper().strip()
            ob_global.gv_dt_filtro_inicio          = "01/01/1990"
            ob_global.gv_dt_filtro_fim             = "01/01/2050"
            ob_global.gv_cc_filtro_empresa         = "TBRA"
            ob_global.gv_cc_gap                    = comum.getParametro('P_CC_GAP').upper().strip()
            ob_global.gv_cc_filtro_outros_filtros  = "TMP.GAP = '" + ob_global.gv_cc_gap + "'"

    except Exception as e:
        v_ds_trace = traceback.format_exc()
        log("ERRO VALIDAÇÃO DOS PARAMETROS DE ENTRADA: " + str(e)+ " >> " + v_ds_trace)
        v_nr_retorno = 93
        
    return v_nr_retorno

def fnc_processar():
    """
        Funcao principal para processar as informacoes
    """
    try:
        v_nr_retorno = 0
        log("-"*150)
        if ob_global.gv_ob_conexao is not None:
            log("conexao               " , "ATIVO"              )           
        log("transacao             " , ob_global.gv_tp_transacao            )
        log("regras                " , ob_global.gv_cc_regras               )     
        log("filtro_inicio         " , ob_global.gv_dt_filtro_inicio        )     
        log("filtro_fim            " , ob_global.gv_dt_filtro_fim           )     
        log("filtro_empresa        " , ob_global.gv_cc_filtro_empresa       )     
        log("filtro_outros_filtros " , ob_global.gv_cc_filtro_outros_filtros)      
        log("filtro_GAP            " , ob_global.gv_cc_gap                  )      
                
        log("-"*150)    
       
        v_cd_erro = ob_global.gv_ob_conexao.var(int)
        v_ds_erro = ob_global.gv_ob_conexao.var(str)

        ob_global.gv_cc_filtro_uf             = '""'
        ob_global.gv_cc_filtro_ie             = '""'
        ob_global.gv_cc_filtro_filial         = '""'
        ob_global.gv_cc_filtro_modelo         = '""'
        ob_global.gv_cc_filtro_serie          = '""'
        ob_global.gv_cc_filtro_nota           = '""'

        v_dt_filtro_inicio = datetime.datetime(int(ob_global.gv_dt_filtro_inicio[6:10]) \
                            , int(ob_global.gv_dt_filtro_inicio[3:5]) \
                            , int(ob_global.gv_dt_filtro_inicio[0:2]))
        v_dt_filtro_fim = datetime.datetime(int(ob_global.gv_dt_filtro_fim[6:10]) \
                            , int(ob_global.gv_dt_filtro_fim[3:5]) \
                            , int(ob_global.gv_dt_filtro_fim[0:2]))
            
            
            
#        log("######     DEBUG     ######")    
#        log("######     DEBUG     ######")    
#        log("######     DEBUG     ######")    
#        log("P_CD_UF=",ob_global.gv_cc_uf)
#        log("P_NM_JOB=",ob_global.gv_nm_job)
#        log("P_CD_ERRO=",v_cd_erro)
#        log("P_DS_ERRO=",v_ds_erro)
#        log("P_DT_FILTRO_INICIO=",v_dt_filtro_inicio)
#        log("P_DT_FILTRO_FIM=",v_dt_filtro_fim)
#        log("P_CC_REGRAS=",ob_global.gv_cc_regras)
#        log("P_CC_FILTRO_UF=",ob_global.gv_cc_filtro_uf)
#        log("P_CC_FILTRO_IE=",ob_global.gv_cc_filtro_ie)
#        log("P_CC_FILTRO_EMPRESA=",ob_global.gv_cc_filtro_empresa)
#        log("P_CC_FILTRO_FILIAL=",ob_global.gv_cc_filtro_filial)
#        log("P_CC_FILTRO_MODELO=",ob_global.gv_cc_filtro_modelo)
#        log("P_CC_FILTRO_SERIE=",ob_global.gv_cc_filtro_serie)
#        log("P_CC_FILTRO_NOTA=",ob_global.gv_cc_filtro_nota)
#        log("P_CC_FILTRO_OUTROS_FILTROS=",ob_global.gv_cc_filtro_outros_filtros)
#        log("P_TP_TRANSACAO=",ob_global.gv_tp_transacao)
#        log("######     DEBUG     ######")    
#        log("######     DEBUG     ######")    
#        log("######     DEBUG     ######")    
            
           

#        ob_global.gv_ob_conexao.executaProcedure("GFCADASTRO.TSH_SANMC_20200_APLICA_REGRAS",
#                                                 P_CD_UF=ob_global.gv_cc_uf,  
#                                                 P_NM_JOB=ob_global.gv_nm_job,   
#                                                 P_CD_ERRO=v_cd_erro,
#                                                 P_DS_ERRO=v_ds_erro,        
#                                                 P_DT_FILTRO_INICIO=v_dt_filtro_inicio,
#                                                 P_DT_FILTRO_FIM=v_dt_filtro_fim,
#                                                 P_CC_REGRAS=ob_global.gv_cc_regras,
#                                                 P_CC_FILTRO_UF=ob_global.gv_cc_filtro_uf,
#                                                 P_CC_FILTRO_IE=ob_global.gv_cc_filtro_ie,
#                                                 P_CC_FILTRO_EMPRESA=ob_global.gv_cc_filtro_empresa,
#                                                 P_CC_FILTRO_FILIAL=ob_global.gv_cc_filtro_filial,
#                                                 P_CC_FILTRO_MODELO=ob_global.gv_cc_filtro_modelo,
#                                                 P_CC_FILTRO_SERIE=ob_global.gv_cc_filtro_serie,
#                                                 P_CC_FILTRO_NOTA=ob_global.gv_cc_filtro_nota,
#                                                 P_CC_FILTRO_OUTROS_FILTROS=ob_global.gv_cc_filtro_outros_filtros,
#                                                 P_TP_TRANSACAO=ob_global.gv_tp_transacao)
       
 
        ob_global.gv_ob_conexao.executaProcedure("GFCADASTRO.TSH_SANMC_20200_APLICA_REGRAS",
                                                 P_NM_JOB=ob_global.gv_nm_job,   
                                                 P_CD_ERRO=v_cd_erro,
                                                 P_DS_ERRO=v_ds_erro,        
                                                 P_DT_FILTRO_INICIO=v_dt_filtro_inicio,
                                                 P_DT_FILTRO_FIM=v_dt_filtro_fim,
                                                 P_CC_REGRAS=ob_global.gv_cc_regras,
                                                 P_CC_FILTRO_OUTROS_FILTROS=ob_global.gv_cc_filtro_outros_filtros,
                                                 P_TP_TRANSACAO=ob_global.gv_tp_transacao)
 
        
        
        
        log(v_cd_erro.getvalue()) 
        log(v_ds_erro.getvalue())
        
        try:
            v_nr_retorno = v_cd_erro.getvalue()
        except:
            v_nr_retorno = 1

        ob_global.gv_ob_conexao.commit()
#        ob_global.gv_ob_conexao.rollback()
        
        return v_nr_retorno
        
    except Exception as e:
        v_ds_trace = traceback.format_exc()
        log("ERRO PROCESSAMENTO: " + str(e)+ " >> " + v_ds_trace)
        v_nr_retorno = 93
        return v_nr_retorno        


if __name__ == "__main__" :
    """
        Ponto de partida
    """    
    # Codigo de Retorno
    v_nr_ret = 0

    # Tratamento de excessao
    v_ds_trace = ''

    # Tratamento de variaveis globais
    ob_global.gv_ob_conexao = None
    # Parametros do arquivo de configuração
    ob_global.gv_tp_transacao = "COMMIT"  
    #ob_global.gv_tp_transacao = "ROLLBACK"
    ob_global.gv_nm_job ="''" 
    ob_global.gv_cc_uf = "FULL"    
    ob_global.gv_cc_regras = "''"                  
    ob_global.gv_dt_filtro_inicio = "01/01/1990"           
    ob_global.gv_dt_filtro_fim = "01/01/2050"              
    ob_global.gv_cc_filtro_empresa = "TBRA"          
    ob_global.gv_tmp_gap = "''"             
    ob_global.gv_cc_filtro_outros_filtros = "TMP.GAP := ‘<Nome do GAP>’"    

    try:

        log("-"*100)
        log("INICIO DA EXECUÇÃO DAS REGRAS PONTUAIS DE MERCADORIA".center(120,'#'))
            
#        log(" ")
        
        # Validacao dos parametros de entrada
        if (v_nr_ret == 0) :
            v_nr_ret = fnc_validar_entrada()
        
#        log(" ")
#        log("Após validacao dos parametros de entrada, v_nr_ret = ",v_nr_ret )
        
        # Verificar conexao com o banco
        if (v_nr_ret == 0) :
            v_nr_ret = fnc_conectar_banco_dados()   
        
#        log(" ")
#        log("Após conexao com o banco, v_nr_ret = ",v_nr_ret )

#        input("Vai processar.... continua?")
        # Processar         
        if (v_nr_ret == 0) :
            v_nr_ret = fnc_processar()                    

        # Finalizacao
#        log(" ")            
        
        if (v_nr_ret == 0) :
            log("SUCESSO NA EXECUÇÃO")
        else:
            log("ERRO NA EXECUÇÃO")
                        
#        log(" ")
    
    except Exception as e:
        v_ds_trace = traceback.format_exc()
        log("ERRO .: " + str(e) + " >> " + v_ds_trace)
        v_nr_ret = 93
    
    sys.exit(v_nr_ret if v_nr_ret >= log.ret else log.ret )

