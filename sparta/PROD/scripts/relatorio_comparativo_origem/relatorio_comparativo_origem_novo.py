#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-

"""
-----------------------------------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: relatorio_tributario.py
  CRIACAO ..: 01/06/2021
  AUTOR ....: Victor Santos Cardoso / KYROS Consultoria
  DESCRICAO : Este relatório tem por finalidade gerar um comparativo entre o Arquivo de Impressão, o registro Original 
              e o Reprocessado ( REGERADO ). Para execução, os parâmetros informamos são: estado (UF), Mês/Ano (MMYYYY), 
              e Série que é opcional, caso não informe a série, o script irá gerar o relatório de todas as séries para 
              o estado e data informada. (Hoje temos uma trava para que seja gerado relatório 
              apenas para as séries '1','C','U T','UK'). 
-----------------------------------------------------------------------------------------------------------------------
  HISTORICO :
    * 01/06/2021 - Victor Santos Cardoso / KYROS Consultoria - Criacao do script.
    * 30/08/2021 - fabrisiag@kyros.com.br 
        PTITES-160 : DESENVOLVIMENTO - RELATÓRIO COMPARATIVO - DV - Inlcusão UF e CFOP no relatório comparativo origens
    * 08/09/2021 - fabrisiag@kyros.com.br
        PTITES-145 : DV - Novo Padrão: INSUMOS - Relatório Comparativo Origens (Impressão x Regerado x Original)
    
    * 08/09/2021 - VICTOR SANTOS CARDOSO - KYROS TECNOLOGIA
        PTITES-145 : ...

    * 08/12/2021 - VICTOR SANTOS CARDOSO - KYROS TECNOLOGIA
        ALTERAÇÃO PARA RECEBER TODAS AS SÉRIES
        ALTERAÇÃO PARA CONVERSÃO DE CARACTERES ESPECIAIS NA PARTE DO PROTOCOLADO
    
    * 23/12/2021 - Welber Pena de Sousa - Kyros Tecnologia
        Alterado o dir_base_r para gravar o relatorio final dentro do diretorio /portaloptrib .

    * 18/03/2022 - Eduardo da Silva Ferreira - Kyros Tecnologia
                 - [PTITES-1719] DV - Controlador por Id Série 


------------------------------------------------------------------------------------------------------------------------
"""
#### PATRONIZACAO PARA O PAINEL DE EXECUCOES....
import sys
import os
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)
import configuracoes
import comum
import util
import sql
comum.log.gerar_log_em_arquivo = True
comum.carregaConfiguracoes(configuracoes)
con=sql.geraCnxBD(configuracoes)

import datetime
from io import FileIO
import re
from openpyxl import Workbook
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.styles.fills import Stop

os.environ['NLS_LANG'] = 'AMERICAN_AMERICA.WE8ISO8859P15'
# os.environ['NLS_LANG'] = 'AMERICAN_AMERICA.AL32UTF8'
configuracoes.var05 = 0 

ret = 0
# global c2
# configuracoes.nome_relatorio = "" 


# REGERADO
# BILLING_COMBINADO

def trucate_table():
    log('Truncando tabelas tab_tmp_impressao e tab_tmp_protocolado')
    query  = """ TRUNCATE TABLE tab_tmp_impressao_fla"""
    query1 = """ TRUNCATE TABLE tab_tmp_protocolado_fla"""

    con.executa(query)
    con.executa(query1)
    con.commit()

def insere_impressao (datai,dataf,filis,ufi):
    trucate_table()
    global c2 
    if configuracoes.banco == 'GFCLONEPREPROD':
        c2 = ''
    else:
        c2 = '@C2'        
    query = """INSERT INTO tab_tmp_impressao_fla
                WITH impressao AS (
                                SELECT
                                L.EMPS_COD                                              EMPS_COD_IMPRESSAO ,
                                L.FILI_COD                                              FILI_COD_IMPRESSAO ,
                                L.SERIE                                                 MNFST_SERIE_IMPRESSAO ,
                                M.MNFST_DTEMISS                                         MNFST_DTEMISS_IMPRESSAO ,
                                M.MNFST_NUM                                             MNFST_NUM_IMPRESSAO ,
                                NULL                                                    TIPO_ASSINANTE_IMPRESSAO ,
                                NULL                                                    TIPO_UTILIZ_IMPRESSAO ,
                                NULL                                                    GRUPO_TENSAO_IMPRESSAO ,
                                M.CADG_COD                                              CADG_COD_IMPRESSAO ,
                                M.CADG_NUM_CONTA                                        TERMINAL_TELEF_IMPRESSAO ,
                                M.CADG_TIP_BILLING                                      IND_CAMPO_01_IMPRESSAO ,
                                DECODE(NVL(M.CADG_TIP, 'X'), 'X', ' ', M.CADG_TIP)      CADG_TIP_IMPRESSAO ,
                                NULL                                                    CADG_TIP_CLI_IMPRESSAO ,
                                NULL                                                    SUB_CLASSE_IMPRESSAO ,
                                NULL                                                    TERMINAL_PRINC_IMPRESSAO ,
                                NULL                                                    CNPJ_EMIT_IMPRESSAO ,                
                                M.CADG_COD_CGCCPF                                       CNPJ_CPF_IMPRESSAO ,
                                NULL                                                    VALOR_CONTABIL_IMPRESSAO ,
                                M.CADG_COD_INSEST                                       IE_IMPRESSAO ,
                                CONVERT(M.CADG_NOM,'us7ascii','WE8ISO8859P15')          RAZAO_SOCIAL_IMPRESSAO ,
                                CONVERT(M.CADG_END,'us7ascii','WE8ISO8859P15')          ENDEREC_IMPRESSAO ,
                                M.CADG_END_NUM_BILLING                                  NUMERO_IMPRESSAO ,
                                CONVERT(M.CADG_END_COMP,'us7ascii','WE8ISO8859P15')     COMPLEMENTO_IMPRESSAO ,
                                M.CADG_END_CEP                                          CEP_IMPRESSAO ,
                                CONVERT(M.CADG_END_BAIRRO,'us7ascii','WE8ISO8859P15')   BAIRRO_IMPRESSAO ,
                                CONVERT(M.CADG_END_MUNIC,'us7ascii','WE8ISO8859P15')    MUNICIPIO_IMPRESSAO ,
                                M.UNFE_SIG                                              UF_IMPRESSAO ,
                                M.CADG_TEL_CONTATO                                      CADG_TEL_CONTATO_IMPRESSAO ,
                                M.CADG_COD                                              CODIDENTCONSUMIDOR_IMPRESSAO ,
                                M.CADG_NUM_CONTA                                        NUMEROTERMINAL_IMPRESSAO ,
                                m.CADG_UF_HABILIT                                       UFHABILITACAO_IMPRESSAO ,
                                M.origem                                                ORIGEM_IMPRESSAO ,
                                M.MIBGE_COD_MUN                                         CODIGOMUNICIPIO_IMPRESSAO ,
                                CASE WHEN nvl(M.CADG_TIP,'X') = 'F' AND VALIDA_DOCUMENTO%s (LPAD(ltrim(M.CADG_COD_CGCCPF,'0'),11,'0')) = 'S' THEN
                                    'S'
                                WHEN nvl(M.CADG_TIP,'X') = 'J' AND VALIDA_DOCUMENTO%s (M.CADG_COD_CGCCPF) = 'S' AND length(M.CADG_COD_CGCCPF) = 14 THEN
                                    'S'
                                ELSE
                                    'N'
                                END                                                     DOCUMENTO_VALIDO_IMPRESSAO ,
                                M.var05                                                 VAR05_IMPRESSAO,
                                CASE WHEN M.FLAG_CAD_DUPLICADO  = 'S' THEN
                                    'SIM'
                                ELSE
                                    'NAO'
                                END                                                     FLAG_CAD_DUPLICADO_IMPRESSAO
                                
                                FROM gfcadastro.billing_combinado_final%s M
                                join gfcarga.tsh_serie_levantamento%s l
                                on l.uf_filial = M.uf_filial
                                    and replace(l.serie,' ','') = M.mnfst_serie
                                    and l.mes_ano = to_date('01' || to_char(M.mnfst_dtemiss,'MMYYYY'),'DDMMYYYY')

                                WHERE
                                replace(l.serie,' ','') = replace(%s,' ','') -- mudança
                                AND m.mnfst_dtemiss >= to_date('%s','DD/MM/YYYY')
                                and m.mnfst_dtemiss <= to_date('%s','DD/MM/YYYY')
                                and m.UF_FILIAL = '%s'
                                --and m.MNFST_NUM >= '000000001'
                                --and m.MNFST_NUM <= '000000020'
                )
                SELECT i.*,
                    (SELECT DECODE(count(1),0,'NAO','SIM')
                        FROM gfcadastro.TBRA_CNPJ%s tc
                        WHERE tc.cnpj = i.CNPJ_CPF_IMPRESSAO
                        AND rownum <= 1) AS CNPJ_TELEFONICA_IMPRESSAO,
                    (SELECT DECODE(count(1),0,'NAO','SIM')
                        FROM gfcadastro.TSH_CPF_AMBIGUO%s tc
                        WHERE tc.CPF_STR = i.CNPJ_CPF_IMPRESSAO
                        AND rownum <= 1) AS CPF_CNPJ_AMBIGUO_IMPRESSAO,

                    sImp.NOME       NOME_IMPRESSAO,
                    bcImp.raiz_cnpj raiz_cnpj_impressao,
                    brIMP.nm_razao_social nm_razao_social_impressao,
                    
                    CASE 
                        WHEN axIMP.DS_SITUACAO_CADASTRAL is null THEN null
                        WHEN UPPER(axIMP.DS_SITUACAO_CADASTRAL)='HABILITADO' THEN axIMP.CC_IE_CONSUMIDOR
                        ELSE 'ISENTO' 
                    END IE_BUREAU_SINTEGRA_IMPRESSAO

                FROM impressao i
                left join GFCADASTRO.TSH_BASE_SERASA%s sIMP
                    ON sIMP.DOCUMENTO = nvl(i.CNPJ_CPF_IMPRESSAO,0)
                    AND sIMP.TIPO      = i.CADG_TIP_IMPRESSAO

                left join GFCADASTRO.TSH_CNPJ_BACEN%s bcIMP
                    ON bcIMP.raiz_cnpj              = SUBSTR(nvl(i.CNPJ_CPF_IMPRESSAO,'XXXXXXXX'),1,8)
                    AND i.DOCUMENTO_VALIDO_IMPRESSAO = 'S'
                    AND length(i.CNPJ_CPF_IMPRESSAO) = 14

                left join GFCADASTRO.TSH_BASE_BUREAU_RECEITA%s brIMP
                    on brIMP.cd_cpfcnpj = I.CNPJ_CPF_IMPRESSAO
                    AND LENGTH(brIMP.nm_razao_social) >= 3
                    AND substr(brIMP.nm_razao_social, 1, 1) <> '0'
                    AND (UPPER(TRIM(nvl(brIMP.nm_razao_social, ' '))) != 'CONSUMIDOR')
                    AND NVL(brIMP.DT_VIGENCIA_SANEAMENTO_INI,I.MNFST_DTEMISS_IMPRESSAO) <= I.MNFST_DTEMISS_IMPRESSAO
                    AND NVL(brIMP.DT_VIGENCIA_SANEAMENTO_FIM,I.MNFST_DTEMISS_IMPRESSAO) >= I.MNFST_DTEMISS_IMPRESSAO
                    AND brIMP.DS_ERRO IS NULL
                    AND (replace(replace(replace(brIMP.nm_razao_social,'.',''),'0',''),' ','')) is NOT null
                                
                left join GFCADASTRO.TSHTB_BUREAU%s axIMP
                    on axIMP.CC_CPFCNPJ_CONSUMIDOR    = I.CNPJ_CPF_IMPRESSAO
                    AND NVL(axIMP.DT_VIGENCIA_SANEAMENTO_INI,I.MNFST_DTEMISS_IMPRESSAO) <= I.MNFST_DTEMISS_IMPRESSAO
                    AND NVL(axIMP.DT_VIGENCIA_SANEAMENTO_FIM,I.MNFST_DTEMISS_IMPRESSAO) >= I.MNFST_DTEMISS_IMPRESSAO
                    AND axIMP.DS_ERRO IS NULL
                    AND TRIM(REPLACE(axIMP.CC_IE_CONSUMIDOR,'0','')) IS NOT NULL

    """%(c2,c2,c2,c2,filis,datai,dataf,ufi,c2,c2,c2,c2,c2,c2)
    # log(query)
    con.executa(query)
    con.commit()
    result = """select count(1) from tab_tmp_impressao_fla"""
    con.executa(result)
    res = con.fetchone()
    log('QUANTIDADE: ', res[0])
    ret = 0
    if result == 0:
        ret = 99
    return(ret)

def insere_protocolado (datai,dataf,filis,ufi):
    global c2 
    if configuracoes.banco == 'GFCLONEPREPROD':
        c2 = ''
    else:
        c2 = '@C2'        
    query = """INSERT INTO tab_tmp_protocolado_fla
                WITH protocolado AS (
                                SELECT
                                l.id_serie_levantamento,
                                l.EMPS_COD                                          EMPS_COD_ORIGINAL ,
                                l.FILI_COD                                          FILI_COD_ORIGINAL ,
                                l.SERIE                                             MNFST_SERIE_ORIGINAL ,
                                mp.DATA_EMISSAO                                     MNFST_DTEMISS_ORIGINAL ,
                                mp.MODELO                                           MODELO_ORIGINAL ,
                                mp.NUMERO_NF                                        MNFST_NUM_ORIGINAL ,
                                MP.CLASSE_CONS                                      TIPO_ASSINANTE_ORIGINAL ,
                                MP.TIPO_UTILIZ                                      TIPO_UTILIZ_ORIGINAL ,
                                MP.GRUPO_TENSAO                                     GRUPO_TENSAO_ORIGINAL ,
                                MP.CADG_COD                                         CADG_COD_ORIGINAL ,
                                MP.TERMINAL_TELEF                                   TERMINAL_TELEF_ORIGINAL ,
                                NVL(DECODE(NVL(TRIM(MP.IND_CAMPO_01), 'X'), '1' , 'J', '2' , 'F', '3', 'E', '4', 'I', MP.IND_CAMPO_01), ' ') ind_campo_01_original,
                                MP.TIPO_CLIENTE                                     CADG_TIP_CLI_ORIGINAL ,
                                MP.SUB_CLASSE                                       SUB_CLASSE_ORIGINAL ,
                                MP.TERMINAL_PRINC                                   TERMINAL_PRINC_ORIGINAL ,
                                MP.CNPJ_EMIT                                        CNPJ_EMIT_ORIGINAL ,
                                D.CNPJ_CPF                                          CNPJ_CPF_ORIGINAL ,
                                MP.VALOR_TOTAL                                      VALOR_CONTABIL_ORIGINAL ,
                                D.IE                                                IE_ORIGINAL ,                                
                                CONVERT(D.RAZAOSOCIAL,'us7ascii','WE8ISO8859P15')   RAZAO_SOCIAL_ORIGINAL ,                                
                                CONVERT(D.ENDERECO,'us7ascii','WE8ISO8859P15')      ENDEREC_ORIGINAL ,
                                TO_CHAR(lpad(D.NUMERO,5,'0'))                       NUMERO_ORIGINAL ,
                                CONVERT(D.COMPLEMENTO,'us7ascii','WE8ISO8859P15')   COMPLEMENTO_ORIGINAL ,
                                lpad(D.CEP, 8, '0')                                 CEP_ORIGINAL ,
                                CONVERT(D.BAIRRO,'us7ascii','WE8ISO8859P15')        BAIRRO_ORIGINAL ,
                                CONVERT(D.MUNICIPIO,'us7ascii','WE8ISO8859P15')     MUNICIPIO_ORIGINAL ,
                                D.UF                                                UF_ORIGINAL ,
                                D.TELEFONECONTATO                                   CADG_TEL_CONTATO_ORIGINAL ,
                                D.CODIDENTCONSUMIDOR                                CODIDENTCONSUMIDOR_ORIGINAL ,
                                D.NUMEROTERMINAL                                    NUMEROTERMINAL_ORIGINAL ,
                                D.UFHABILITACAO                                     UFHABILITACAO_ORIGINAL ,
                                D.CODIGOMUNICIPIO                                   CODIGOMUNICIPIO_ORIGINAL ,
                                VALIDA_DOCUMENTO%s(D.CNPJ_CPF)                      DOCUMENTO_VALIDO_ORIGINAL , 
                                CASE WHEN to_char(mp.data_emissao, 'YYYY') <= 2016 THEN
                                    CASE WHEN valida_documento%s(mp.cnpj_cpf) = 'S' AND length(mp.cnpj_cpf) = 14 THEN
                                        'J'
                                    ELSE
                                        'F'
                                    END
                                ELSE
                                    decode(mp.ind_campo_01, 1, 'J', 2, 'F', 3, 'E', 4, 'I', 'F')
                                END                                                 CADG_TIP_ORIGINAL
                                , I.UF AS                                           UF_ITEM_ORIGINAL 
                                , I.CFOP AS                                         CFOP_ITEM_ORIGINAL 
                                , LPAD(TRIM(I.NUM_ITEM),5,'0') AS                   NUM_ITEM_ORIGINAL 
                                , I.NUM_ITEM 

                                from GFCARGA.TSH_MESTRE_CONV_115%s MP                                
                                JOIN GFCARGA.TSH_ITEM_CONV_115%s I
                                ON I.rowid = (SELECT MIN(I3.rowid) KEEP (DENSE_RANK FIRST ORDER BY CASE I3.CFOP WHEN '0000' THEN 2 ELSE 1 END, I3.NUM_ITEM)
                                                FROM GFCARGA.TSH_ITEM_CONV_115%s I3
                                                WHERE I3.ID_SERIE_LEVANTAMENTO = MP.ID_SERIE_LEVANTAMENTO
                                                AND I3.UF_FILIAL = MP.UF_FILIAL
                                                AND I3.NUMERO_NF = MP.NUMERO_NF)
                                -- FIM
                                join GFCARGA.TSH_DESTINATARIO_CONV_115%s D
                                on MP.ID_SERIE_LEVANTAMENTO = D.ID_SERIE_LEVANTAMENTO
                                AND MP.VOLUME = D.VOLUME
                                AND MP.LINHA = D.LINHA
                                AND MP.UF_FILIAL = D.UF_FILIAL

                                join GFCARGA.TSH_SERIE_LEVANTAMENTO%s L
                                on MP.ID_SERIE_LEVANTAMENTO = L.ID_SERIE_LEVANTAMENTO

                                WHERE replace(l.SERIE,' ','') = replace(%s ,' ','')
                                AND L.mes_ano = to_date('%s','DD/MM/YYYY')
                                and MP.UF_FILIAL = '%s'
                                --and MP.NUMERO_NF >= '000000001'
                                --and MP.NUMERO_NF <= '000000020'
                )
                SELECT p.*,
                    (SELECT DECODE(count(1),0,'NAO','SIM')
                        FROM gfcadastro.TBRA_CNPJ%s tc
                        WHERE tc.cnpj = CNPJ_CPF_ORIGINAL
                        AND rownum <= 1) CNPJ_TELEFONICA_ORIGINAL,

                    sOri.NOME             NOME_ORIGINAL,
                    bcOri.raiz_cnpj       raiz_cnpj_original,
                    brOri.nm_razao_social nm_razao_social_original,
                    
                    CASE 
                        WHEN axOri.DS_SITUACAO_CADASTRAL is null THEN null
                        WHEN UPPER(axOri.DS_SITUACAO_CADASTRAL)='HABILITADO' THEN axOri.CC_IE_CONSUMIDOR
                        ELSE 'ISENTO' 
                    END IE_BUREAU_SINTEGRA_ORIGINAL

                FROM protocolado p

                left join GFCADASTRO.TSH_BASE_SERASA%s sORI
                    on sORI.DOCUMENTO = nvl(p.CNPJ_CPF_ORIGINAL,0)
                    AND sORI.TIPO      = p.CADG_TIP_ORIGINAL

                left join GFCADASTRO.TSH_CNPJ_BACEN%s bcORI
                    on bcORI.raiz_cnpj             = SUBSTR(nvl(p.CNPJ_CPF_ORIGINAL,'XXXXXXXX'),1,8)
                    AND p.DOCUMENTO_VALIDO_ORIGINAL = 'S'
                    AND length(p.CNPJ_CPF_ORIGINAL) = 14
                            
                left join GFCADASTRO.TSH_BASE_BUREAU_RECEITA%s brORI
                    on brORI.cd_cpfcnpj = p.CNPJ_CPF_ORIGINAL
                    AND LENGTH(brORI.nm_razao_social) >= 3
                    AND substr(brORI.nm_razao_social, 1, 1) <> '0'
                    AND (UPPER(TRIM(nvl(brORI.nm_razao_social, ' '))) != 'CONSUMIDOR')
                    AND NVL(brORI.DT_VIGENCIA_SANEAMENTO_INI,p.MNFST_DTEMISS_ORIGINAL) <= p.MNFST_DTEMISS_ORIGINAL
                    AND NVL(brORI.DT_VIGENCIA_SANEAMENTO_FIM,p.MNFST_DTEMISS_ORIGINAL) >= p.MNFST_DTEMISS_ORIGINAL
                    AND brORI.DS_ERRO IS NULL
                    AND (replace(replace(replace(brORI.nm_razao_social,'.',''),'0',''),' ','')) is NOT null
                            
                left join GFCADASTRO.TSHTB_BUREAU%s axORI
                    on     axORI.CC_CPFCNPJ_CONSUMIDOR    = p.CNPJ_CPF_ORIGINAL
                    AND NVL(axORI.DT_VIGENCIA_SANEAMENTO_INI,p.MNFST_DTEMISS_ORIGINAL) <= p.MNFST_DTEMISS_ORIGINAL
                    AND NVL(axORI.DT_VIGENCIA_SANEAMENTO_FIM,p.MNFST_DTEMISS_ORIGINAL) >= p.MNFST_DTEMISS_ORIGINAL
                    AND axORI.DS_ERRO IS NULL
                    AND TRIM(REPLACE(axORI.CC_IE_CONSUMIDOR,'0','')) IS NOT NULL

    """%(c2,c2,c2,c2,c2,c2,c2,filis,datai,ufi,c2,c2,c2,c2,c2)
    # log(query)
    con.executa(query)
    con.commit()
    result = """select count(1) from tab_tmp_protocolado_fla"""
    con.executa(result)
    res = con.fetchone()
    log('QUANTIDADE: ', res[0])
    ret = 0
    if result == 0:
        ret = 99
    return(ret)

def busca_relatorio (datai,dataf,filis,ufi):
    global c2 
    if configuracoes.banco == 'GFCLONEPREPROD':
        c2 = ''
    else:
        c2 = '@C2'  

    query = """
           WITH
                -- MUDANÇA TEIXEIRA 21/10/2021 - DEV VICTOR SANTOS 
                ITEM as (
                select 
                    i.EMPS_COD, 
                    i.FILI_COD, 
                    i.INFST_SERIE, 
                    i.INFST_DTEMISS, 
                    i.INFST_NUM,             
                    I.UF AS UF_ITEM_REGERADO,
                    I.CFOP AS CFOP_ITEM_REGERADO,
                    LPAD(TRIM(TO_CHAR(I.INFST_NUM_SEQ)),5,'0') NUM_ITEM_REGERADO,
                    RANK() OVER (PARTITION BY i.EMPS_COD, i.FILI_COD, i.INFST_SERIE, i.INFST_DTEMISS, i.INFST_NUM order by CFOP desc, infst_num_seq) LINHA
                from openrisow.ITEM_NFTL_SERV i
                join gfcarga.TSH_SERIE_LEVANTAMENTO l
                    on l.emps_cod = i.emps_cod
                and l.fili_cod = i.fili_cod
                and l.serie    = i.infst_serie
                and i.infst_dtemiss between l.mes_ano and last_day(l.mes_ano)
                WHERE replace(l.serie,' ','') = replace(%s,' ','') -- mudança
                AND infst_dtemiss >= to_date('%s','DD/MM/YYYY')
                and infst_dtemiss <= to_date('%s','DD/MM/YYYY')
                and l.uf_filial = '%s' 
                --and i.INFST_NUM >= '000000001' 
                --and i.INFST_NUM <= '000000020'
                ) ,
                REGERADO as
                (SELECT
                M.EMPS_COD                                                                      EMPS_COD_REGERADO,
                M.FILI_COD                                                                      FILI_COD_REGERADO,
                M.MNFST_SERIE                                                                   MNFST_SERIE_REGERADO,
                M.MNFST_DTEMISS                                                                 MNFST_DTEMISS_REGERADO, 
                M.MNFST_NUM                                                                     MNFST_NUM_REGERADO,
                M.MNFST_VAL_TOT                                                                 VALOR_CONTABIL_REGERADO,
                NVL(TRIM(comp.CADG_TIP_ASSIN), '0')                                             TIPO_ASSINANTE_REGERADO,
                TRIM(CASE 
                        WHEN NVL(M.MNFST_TIP_UTIL, '0') = '0' OR M.MNFST_TIP_UTIL = '0' THEN
                            MNFST_TIP_UTIL
                    ELSE
                        NVL(CADG_TIP_UTILIZ, '1')
                END)                                                                            TIPO_UTILIZ_REGERADO,
                NVL(TRIM(comp.CADG_GRP_TENSAO), '00')                                           GRUPO_TENSAO_REGERADO,
                nvl(SUBSTR(M.CADG_COD, 1, 12), ' ')                                             CADG_COD_REGERADO,
                nvl(comp.CADG_NUM_CONTA, ' ')                                                   TERMINAL_TELEF_REGERADO,
                cad.CADG_TIP                                                                    CADG_TIP_REGERADO,            
                NVL(comp.CADG_TIP_CLI, ' ')                                                     CADG_TIP_CLI_REGERADO,
                NVL(comp.CADG_SUB_CONSU, '0')                                                   SUB_CLASSE_REGERADO,
                --T.TERMINAL_PRINCIPAL                                                          TERMINAL_PRINC_REGERADO,
                lpad(cad.CADG_COD_CGCCPF,14,'0')                                                CNPJ_CPF_REGERADO,
                nvl(cad.CADG_COD_INSEST, ' ')                                                   IE_REGERADO,
                nvl(SUBSTR(CONVERT(cad.CADG_NOM,'us7ascii','WE8ISO8859P15'), 1, 35), ' ')       RAZAO_SOCIAL_REGERADO,                
                nvl(SUBSTR(CONVERT(cad.CADG_END,'us7ascii','WE8ISO8859P15'), 1, 45), ' ')       ENDEREC_REGERADO,
                nvl(lpad(cad.CADG_END_NUM,5,'0'), 0)                                            NUMERO_REGERADO,
                nvl(SUBSTR(CONVERT(cad.CADG_END_COMP,'us7ascii','WE8ISO8859P15'), 1, 15), ' ')  COMPLEMENTO_REGERADO,
                nvl(cad.CADG_END_CEP, '0')                                                      CEP_REGERADO,
                nvl(CONVERT(cad.CADG_END_BAIRRO,'us7ascii','WE8ISO8859P15'), ' ')               BAIRRO_REGERADO,
                nvl(CONVERT(cad.CADG_END_MUNIC,'us7ascii','WE8ISO8859P15'), ' ')                MUNICIPIO_REGERADO,
                nvl(cad.UNFE_SIG, ' ')                                                          UF_REGERADO,
                nvl(comp.CADG_TEL_CONTATO, ' ')                                                 CADG_TEL_CONTATO_REGERADO,
                nvl(SUBSTR(cad.CADG_COD, 1, 12), ' ')                                           CODIDENTCONSUMIDOR_REGERADO,
                nvl(comp.CADG_NUM_CONTA, ' ')                                                   NUMEROTERMINAL_REGERADO,
                nvl(comp.CADG_UF_HABILIT, ' ')                                                  UFHABILITACAO_REGERADO,
                nvl(cad.MIBGE_COD_MUN, 0)                                                       CODIGOMUNICIPIO_REGERADO,
                replace(replace(replace(replace(CAD.VAR05
                ,'TESHUVA_CARGA_BILLING' , '')
                ,'TESHUVA_CARGA_PROTOCOLADO', '')
                ,' (IMPR)' , '')
                ,' (PROT)' , '')                                                                VAR05_REGERADO
                , UF_ITEM_REGERADO
                , CFOP_ITEM_REGERADO
                , NUM_ITEM_REGERADO
                
                from OPENRISOW.MESTRE_NFTL_SERV M
                left JOIN ITEM I
                on I.INFST_NUM = M.MNFST_NUM
                AND I.INFST_DTEMISS = M.MNFST_DTEMISS
                AND I.INFST_SERIE = M.MNFST_SERIE
                AND I.EMPS_COD = M.EMPS_COD
                AND I.FILI_COD = M.FILI_COD
                and i.linha = 1

                join gfcarga.TSH_SERIE_LEVANTAMENTO l
                on l.emps_cod = m.emps_cod
                and l.fili_cod = m.fili_cod
                and l.serie = m.mnfst_serie
                and m.mnfst_dtemiss between l.mes_ano and last_day(l.mes_ano)
                join openrisow.CLI_FORNEC_TRANSP cad
                on m.cadg_cod = cad.cadg_cod
                AND m.catg_cod = cad.catg_cod
                AND cad.cadg_dat_atua = (SELECT max(t.cadg_dat_atua)
                from openrisow.CLI_FORNEC_TRANSP t
                where t.cadg_cod = m.cadg_cod
                and t.catg_cod = m.catg_cod
                and t.cadg_dat_atua <= m.mnfst_dtemiss)
                join openrisow.COMPLVU_CLIFORNEC comp
                on comp.cadg_cod = cad.cadg_cod
                AND comp.catg_cod = cad.catg_cod
                AND comp.cadg_dat_atua = cad.cadg_dat_atua
                WHERE replace(l.serie,' ','') = replace(%s,' ','') -- mudança
                AND m.mnfst_dtemiss >= to_date('%s','DD/MM/YYYY')
                and m.mnfst_dtemiss <= to_date('%s','DD/MM/YYYY')
                and l.uf_filial = '%s'
                --and M.MNFST_NUM >= '000000001' 
                --and M.MNFST_NUM <= '000000020'
                )

                SELECT 
                FILI_COD_CGC , FILI_COD_INSEST , MNFST_NUM_ORIGINAL,
                MNFST_SERIE_ORIGINAL, MODELO_ORIGINAL ,
                p.NUM_ITEM ,
                ' ' AS UF_ITEM_IMPRESSAO , NVL(p.UF_ITEM_ORIGINAL,' ') AS UF_ITEM_ORIGINAL , NVL(r.UF_ITEM_REGERADO, ' ') AS UF_ITEM_REGERADO ,
                ' ' AS CFOP_ITEM_IMPRESSAO , NVL(p.CFOP_ITEM_ORIGINAL,' ') AS CFOP_ITEM_ORIGINAL , NVL(r.CFOP_ITEM_REGERADO, ' ') AS CFOP_ITEM_REGERADO ,
                                
                CASE 
                  WHEN DOCUMENTO_VALIDO_IMPRESSAO = 'S' THEN
                    CNPJ_TELEFONICA_IMPRESSAO
                ELSE
                   CNPJ_TELEFONICA_ORIGINAL
                END AS CNPJ_TELEFONICA ,
                
                CASE 
                  WHEN DOCUMENTO_VALIDO_IMPRESSAO = 'S' THEN
                    CPF_CNPJ_AMBIGUO_IMPRESSAO
                ELSE
                    'NAO'
                END AS CPF_CNPJ_AMBIGUO,

                CNPJ_CPF_IMPRESSAO         , CNPJ_CPF_ORIGINAL            , CNPJ_CPF_REGERADO ,
                
                CADG_TIP_IMPRESSAO           , CADG_TIP_ORIGINAL            , CADG_TIP_REGERADO ,
                VALOR_CONTABIL_IMPRESSAO     , VALOR_CONTABIL_ORIGINAL      , VALOR_CONTABIL_REGERADO , 
                
                DECODE(DECODE(DOCUMENTO_VALIDO_IMPRESSAO,'S',i.raiz_cnpj_impressao,p.raiz_cnpj_original),null,null,'ISENTO') IE_BACEN,
                DECODE(DOCUMENTO_VALIDO_IMPRESSAO,'S', i.IE_BUREAU_SINTEGRA_IMPRESSAO ,p.IE_BUREAU_SINTEGRA_ORIGINAL) AS IE_BAUREU_SINTEGRA,
                IE_IMPRESSAO                 , IE_ORIGINAL                  , IE_REGERADO ,
                
                DECODE(DOCUMENTO_VALIDO_IMPRESSAO,'S', i.nm_razao_social_impressao, p.nm_razao_social_original) as RAZAO_SOCIAL_RECEITA_FEDERAL,
                decode(DOCUMENTO_VALIDO_IMPRESSAO,'S',i.NOME_IMPRESSAO,p.NOME_ORIGINAL) RAZAO_SOCIAL_BUREAU_SERASA,
                
                RAZAO_SOCIAL_IMPRESSAO       , RAZAO_SOCIAL_ORIGINAL        , RAZAO_SOCIAL_REGERADO ,
                ENDEREC_IMPRESSAO            , ENDEREC_ORIGINAL             , ENDEREC_REGERADO ,
                NUMERO_IMPRESSAO             , NUMERO_ORIGINAL              , NUMERO_REGERADO ,
                COMPLEMENTO_IMPRESSAO        , COMPLEMENTO_ORIGINAL         , COMPLEMENTO_REGERADO ,
                CEP_IMPRESSAO                , CEP_ORIGINAL                 , CEP_REGERADO ,
                BAIRRO_IMPRESSAO             , BAIRRO_ORIGINAL              , BAIRRO_REGERADO ,
                MUNICIPIO_IMPRESSAO          , MUNICIPIO_ORIGINAL           , MUNICIPIO_REGERADO ,
                UF_IMPRESSAO                 , UF_ORIGINAL                  , UF_REGERADO    ,
                CADG_TEL_CONTATO_IMPRESSAO   , CADG_TEL_CONTATO_ORIGINAL    , CADG_TEL_CONTATO_REGERADO ,
                CODIDENTCONSUMIDOR_IMPRESSAO , CODIDENTCONSUMIDOR_ORIGINAL  , CODIDENTCONSUMIDOR_REGERADO ,
                NUMEROTERMINAL_IMPRESSAO     , NUMEROTERMINAL_ORIGINAL      , NUMEROTERMINAL_REGERADO ,
                UFHABILITACAO_IMPRESSAO      , UFHABILITACAO_ORIGINAL       , UFHABILITACAO_REGERADO ,
                MNFST_DTEMISS_IMPRESSAO      , MNFST_DTEMISS_ORIGINAL       , MNFST_DTEMISS_REGERADO ,
                CODIGOMUNICIPIO_IMPRESSAO    , CODIGOMUNICIPIO_ORIGINAL     , CODIGOMUNICIPIO_REGERADO,
                nvl(ORIGEM_IMPRESSAO,'CONV115_ORIGINAL')                    , FLAG_CAD_DUPLICADO_IMPRESSAO ,
                VAR05_REGERADO             , VAR05_IMPRESSAO            
                        
                from tab_tmp_protocolado_fla p
                
                left join tab_tmp_impressao_fla I
                ON p.EMPS_COD_ORIGINAL       = I.EMPS_COD_IMPRESSAO
                AND p.FILI_COD_ORIGINAL      = I.FILI_COD_IMPRESSAO
                AND p.MNFST_SERIE_ORIGINAL   = I.MNFST_SERIE_IMPRESSAO
                AND p.MNFST_NUM_ORIGINAL     = I.MNFST_NUM_IMPRESSAO
                AND p.MNFST_DTEMISS_ORIGINAL = I.MNFST_DTEMISS_IMPRESSAO
                
                left join REGERADO r
                ON p.EMPS_COD_ORIGINAL       = R.EMPS_COD_REGERADO
                AND p.FILI_COD_ORIGINAL      = R.FILI_COD_REGERADO
                AND p.MNFST_SERIE_ORIGINAL   = R.MNFST_SERIE_REGERADO
                AND p.MNFST_DTEMISS_ORIGINAL = R.MNFST_DTEMISS_REGERADO
                AND p.MNFST_NUM_ORIGINAL     = R.MNFST_NUM_REGERADO                
                  
                left join openrisow.FILIAL c
                on p.FILI_COD_ORIGINAL = c.FILI_COD
                and p.EMPS_COD_ORIGINAL = C.EMPS_COD
                order by p.MNFST_NUM_ORIGINAL
    """%(filis,datai,dataf,ufi,filis,datai,dataf,ufi)
    # log(query)
    retorno = [[]]
    lin = 0 
    retorno[0]=[
        "CNPJ Telefonica",  
        "IE Telefonica",   
        "Nº NF",           
        "SERIE",           
        "MODELO",
        
		"Sequencial Item", 

        "UF Item_Arq. Imp",
        "UF Item_Conv. 115 Orig",
        "UF Item_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",
        
        "CFOP Item_Arq. Imp",
        "CFOP Item_Conv. 115 Orig",
        "CFOP Item_Conv. 115 Rep",        
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep", 

         # ALTERAÇÃO VICTOR
        "CNPJ TELEFONICA",
        "CPF Ambíguo",          
        "CNPJ ou CPF_Arq. Imp.",
        "CNPJ ou CPF_Conv. 115 Orig",
        "CNPJ ou CPF_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",
        
        # ALTERAÇÃO VICTOR
        "Tipo_de_Cliente_Arq. Imp.", 
        "Tipo_de_Cliente_Conv. 115 Orig",
        "Tipo_de_Cliente_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "Valor_Contabil_Arq. Imp.", 
        "Valor_Contabil_Conv. 115 Orig",
        "Valor_Contabil_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "IE_BACEN",
        "IE_Bureau Sintegra",        
        "IE_Arq. Imp.",
        "IE_Conv. 115 Orig",
        "IE_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",        
        
        "Razao Social Receita Federal",
        "Razao Social Bureau Serasa",        
        "Razao Social_Arq. Imp.",
        "Razao Social_Conv. 115 Orig",
        "Razao Social_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "Logradouro_Arq. Imp.",
        "Logradouro_Conv. 115 Orig",
        "Logradouro_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "Numero_Arq. Imp.",
        "Numero_Conv. 115 Orig",
        "Numero_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "Complemento_Arq. Imp.",
        "Complemento_Conv. 115 Orig",
        "Complemento_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "CEP_Arq. Imp.",
        "CEP_Conv. 115 Orig",
        "CEP_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "Bairro_Arq. Imp.",
        "Bairro_Conv. 115 Orig",
        "Bairro_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "Municipio_Arq. Imp.",
        "Municipio_Conv. 115 Orig",
        "Municipio_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "UF_Arq. Imp.",
        "UF_Conv. 115 Orig",
        "UF_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",
        
        "Telefone de contato_Arq. Imp.",
        "Telefone de contato_Conv. 115 Orig",
        "Telefone de contato_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "Código de identificação do consumidor_Arq. Imp.",
        "Código de identificação do consumidor_Conv. 115 Orig",
        "Código de identificação do consumidor_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "Numero do terminal telefônico_Arq. Imp.",
        "Numero do terminal telefônico_Conv. 115 Orig",
        "Numero do terminal telefônico_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "UF de habilitação_Arq. Imp.",
        "UF de habilitação_Conv. 115 Orig",
        "UF de habilitação_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "Data de emissao_Arq. Imp.",
        "Data de emissao_Conv. 115 Orig",
        "Data de emissao_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "Código do Municipio_Arq. Imp.",
        "Código do Municipio_Conv. 115 Orig",
        "Código do Municipio_Conv. 115 Rep",
        "Comparativo Arq. Imp. x c115 rep",
        "Comparativo c115 orig x c115 rep",

        "Origem",
        "FLAG de cadastro duplicado",
        "Regra Aplicada 115 Rep"
        
    ]
    
    # log(query)
    con.executa(query)
    # result = con.fetchall()
    result = con.fetchone()
    
    colunas = con.description()
    
    configuracoes.var05

    for x in range(len(colunas)):
        if colunas[x][0] == 'VAR05_IMPRESSAO':
            configuracoes.var05 = x

    if not result:
        log("#### ATENÇÃO: Nenhum Resultado para query")
        log("####     Query = ")
        log("####")
        log(query)
        log("####")
        ret=99
        return(ret)
    else:
        # retorno += result
        ant = ''
        try:            
            while result:
                ant = result
                retorno.append(result)
                result = con.fetchone()
        except Exception as e:
            log('#### ERRO, NOTA ANTERIOR AO ERRO: ', ant)
            log('ERRO - ' , e)
            ret=99
            return(ret)

    return(retorno)
          
def processar(p_id_serie):
    ret              = 0
    dadosSerie       = comum.buscaDadosSerie(p_id_serie)  
    if not dadosSerie:
        log("#### ATENÇÃO: Nenhum Resultado para busca de séries")
        log("#### - Relatório não gerado ", filiais)
        log("####")
        ret = 99
        return ret
    
    ufi              = dadosSerie['uf']
    mesanoi          = str(dadosSerie['mes']) + str(dadosSerie['ano'])
    mesi             = mesanoi[:2]
    anoi             = mesanoi[2:]
    seriesi          = dadosSerie['serie']
    diretorio        = dadosSerie['dir_serie']
        
    filiali          = dadosSerie['filial']
    id_seriei        = dadosSerie['id_serie']
    configuracoes.var05  
   
    anoR    = mesanoi[4:]
    datai   = "01/" + mesi + "/" + anoi
    dataf   =  str(ultimodia(int(anoi),int(mesi)))+"/"+str(mesi)+"/"+str(anoi)
    
    log('SÉRIE       = ' , str(dadosSerie['serie']))
    log('FILIAL      = ' , str(dadosSerie['filial']))
    log('IE          = ' , str(dadosSerie['id_serie']))
    log('DATA INICIO = ' , datai)
    log('DATA FIM    = ' , dataf)
    
    log("# Iniciando busca de dados para o relatório com a série " + str(dadosSerie['serie']))

    #### Monta caminho e nome do relatorio MUDAR PORTALOPTRB
    dir_relatorio       = os.path.join(diretorio, 'INSUMOS')
    log('DIRETORIO      = ' , dir_relatorio)
    
    #### Se a pasta do relatório não existir, cria
    if not os.path.isdir(dir_relatorio) :
        os.makedirs(dir_relatorio)
        
    log("*" *100)
    log("# Iniciando a inserção de dados na tabela TAB_TMP_IMPRESSAO")
    impressao = insere_impressao(datai,dataf,"'"+str(dadosSerie['serie'])+"'",ufi)
    if impressao > 0:
        log("####")
        log("#### ATENÇÃO: Nenhum Resultado para o insert na tabela TAB_TMP_IMPRESSAO")
        log("####")
    else:
        log("# SUCESSO!!! Dados inseridos na tabela TAB_TMP_IMPRESSAO")
        log("*" *100)

    log("# Iniciando a inserção de dados na tabela TAB_TMP_PROTOCOLADO")
    protocolado = insere_protocolado(datai,dataf,"'"+str(dadosSerie['serie'])+"'",ufi)
    if protocolado > 0:
        log("####")
        log("#### ATENÇÃO: Nenhum Resultado para o insert na tabela TAB_TMP_PROTOCOLADO")
        log("####")
    else:
        log("# SUCESSO!!! Dados inseridos na tabela TAB_TMP_PROTOCOLADO")
        log("*" *100)
    
    log("# Iniciando a consulta para criação do relatório...")
    dados_relatorio = busca_relatorio(datai,dataf,"'"+str(dadosSerie['serie'])+"'",ufi)
    
    if type(dados_relatorio) != list:
        return(90)                

    log("#    Fim da consulta... Gerando relatório.")
    log("*" *100) 

    dicionario = {}
    """
        COLUNA 1 - VALOR DO ITEM IMPRESSÃO
        COLUNA 2 - VALOR DO ITEM ORIGINAL
        COLUNA 3 - VALOR DO ITEM REPROCESSADO
        COLUNA 4 - CONTADOR CRÍTICAS IMPRESSÃO X REPROCESSADO
        COLUNA 5 - CONTADOR CRÍTICAS ORIGINAL X REPROCESSADO
        COLUNA 6 - COLUNA DIFERENTE OU IGUAL IMPRESSÃO X REPROCESSADO
        COLUNA 7 - COLUNA DIFERENTE OU IGUAL ORIGINAL X REPROCESSADO
        COLUNA 8 - Campos inexistentes no layout de emissão de 2016 PRA TRÁS
        COLUNA 9 - No código de cliente se direferente e realizando trim e lpad de 12 zeros a esquerda 
                   ficar igual então classificar como DIFERENTE – APENAS ZEROS A ESQUERDA
        COLUNA 10 - CONTADOR PARA OS CASOS:  APENAS ZERO A ESQUERDA IMPRESSÃO X REPROCESSADO
        COLUNA 11 - CONTADOR PARA OS CASOS:  APENAS ZERO A ESQUERDA ORIGINAL X REPROCESSADO
        COLUNA 12 - CONVERSÃO DE DATA
    """
    # 
    # INICIO
    # dicionario ["UF Item"]         = [6,7,8,0,0,9,10,False,False,0,0, False]
    # dicionario ["CFOP_ITEM"]       = [11,12,13,0,0,14,15,False,False,0,0,False]
    # #FIM
    # dicionario ["CPF"]             = [16,17,18,0,0,19,20,False,False,0,0,False]
    # # VICTOR ALTERAÇÃO 
    # dicionario ["TIPO_CLI"]        = [21,22,23,0,0,24,25,False,False,0,0,False]
    # dicionario ["VALOR_CONTABIL"]  = [26,27,28,0,0,29,30,False,False,0,0,False]
    # # FIM 
    # dicionario ["IE"]              = [31,32,33,0,0,36,37,False,False,0,0,False]
    # dicionario ["RAZAO_SOCIAL"]    = [38,39,40,0,0,42,43,False,False,0,0,False]
    # dicionario ["LOGRADOURO"]      = [44,45,46,0,0,47,48,False,False,0,0,False]
    # dicionario ["NUMERO"]          = [49,50,51,0,0,52,53,False,False,0,0,False]
    # dicionario ["COMPLEMENTO"]     = [54,55,56,0,0,57,58,False,False,0,0,False]
    # dicionario ["CEP"]             = [59,60,61,0,0,62,63,False,False,0,0,False]
    # dicionario ["BAIRRO"]          = [64,65,66,0,0,67,68,False,False,0,0,False]
    # dicionario ["MUNICIPIO"]       = [69,70,71,0,0,72,73,False,False,0,0,False]
    # dicionario ["UF"]              = [74,75,76,0,0,77,78,False,False,0,0,False]
    # dicionario ["TEL_CONTATO"]     = [79,80,81,0,0,82,83,False, True,0,0,False]
    # dicionario ["COD_CONSUMIDOR"]  = [84,85,86,0,0,87,88,False, True,0,0,False]
    # dicionario ["TERMINAL"]        = [89,90,91,0,0,92,93,False,False,0,0,False]
    # dicionario ["UF_HABIT"]        = [94,95,96,0,0,97,98,False,False,0,0,False]
    # dicionario ["DT_EMISS"]        = [99,100,101,0,0,102,103,False,False,0,0, True]
    # dicionario ["COD_MUNIC"]       = [104,105,106,0,0,107,108,True,False,0,0,False]

    dicionario ["UF Item"]                                = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["CFOP Item"]                              = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["CNPJ ou CPF"]                            = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["Tipo_de_Cliente"]                        = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["Valor_Contabil"]                         = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["IE"]                                     = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["Razao Social"]                           = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["Logradouro"]                             = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["Numero"]                                 = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["Complemento"]                            = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["CEP"]                                    = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["Bairro"]                                 = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["Municipio"]                              = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["UF"]                                     = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["Telefone de contato"]                    = [0,0,0,0,0,0,0 ,False, True,0,0,False]
    dicionario ["Código de identificação do consumidor"]  = [0,0,0,0,0,0,0 ,False, True,0,0,False]
    dicionario ["Numero do terminal telefônico"]          = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["UF de habilitação"]                      = [0,0,0,0,0,0,0 ,False,False,0,0,False]
    dicionario ["Data de emissao"]                        = [0,0,0,0,0,0,0 ,False,False,0,0, True]
    dicionario ["Código do Municipio"]                    = [0,0,0,0,0,0,0 , True,False,0,0,False]

    for k in dicionario.keys():
        for c in dados_relatorio[0]:
            if dicionario[k][0] == 0:
                if c.startswith('%s_Arq'%k):
                    col = dados_relatorio[0].index(c)
                    dicionario[k][0] = col
                    dicionario[k][1] = col + 1
                    dicionario[k][2] = col + 2
                    dicionario[k][5] = col + 3
                    dicionario[k][6] = col + 4
    count  = 0
    volume = 0 
    relatorio = False   
    dicionario_keys = dicionario.keys()

    for row in dados_relatorio[1:]: 
        
        relatar = False
        
        if not relatorio or count == 500000:
            
            count = 0
            
            if relatorio:
                relatorio.close()
                os.system('rm ' + nome_relatorio + '.zip')
                dir_corrente = os.getcwd()
                os.chdir(dir_relatorio)                    
                compacta = 'zip -q -n zip ' + nome_arquivo + '.zip ' + nome_arquivo  
                os.system(compacta)
                os.system('rm ' + nome_relatorio )
                os.chdir(dir_corrente)
    
                volume += 1     

            nome_arquivo = "Comparativo_Conv115_Original_x_Impressao_x_Regerado_"+ufi+"_"+str(dadosSerie['serie']).replace(' ', '')+"_"+mesanoi+"_"+str(volume).rjust(3,'0')+".csv"
            nome_relatorio = os.path.join(dir_relatorio, nome_arquivo)
            relatorio = open(nome_relatorio,'w',encoding='iso-8859-1', errors='replace')

            relatorio.write(';'.join(str(x) for x in dados_relatorio[0]) + '\n')

        linha = [x if x else '' for x in row] 
        
        # FIXO NO CÓDIGO, PARAMETRO DE LIMITE DE 10000
        flagLimite = 'N'

        for critica in dicionario_keys:

            linha.insert(dicionario[critica][5],'')
            linha.insert(dicionario[critica][6],'')

        if linha[configuracoes.var05].__contains__('|R2;>>'): # qualquer lugar 
            dt = linha[configuracoes.var05].split('R2;>>')[-1].split('->')[0].split('/')                
            data = datetime.datetime(2000+int(dt[2]), int(dt[1]), int(dt[0]))
            linha[dicionario["Data de emissao"][0]] = data  

        for critica in dicionario_keys:

            if flagLimite == 'N' or (dicionario[critica][3] < 10000 and dicionario[critica][4] < 10000): 

                flagCritica = True

                if dicionario[critica][7]:
                    if linha[dicionario["Data de emissao"][1]].year <= 2016:
                        linha[dicionario[critica][0]] = 'Campo Inexistente em ' + str(linha[dicionario["Data de emissao"][1]].year)           
                        linha[dicionario[critica][1]] = 'Campo Inexistente em ' + str(linha[dicionario["Data de emissao"][1]].year)  
                        linha[dicionario[critica][2]] = 'Campo Inexistente em ' + str(linha[dicionario["Data de emissao"][1]].year)  
                        linha[dicionario[critica][5]] = 'Campo Inexistente em ' + str(linha[dicionario["Data de emissao"][1]].year)  
                        linha[dicionario[critica][6]] = 'Campo Inexistente em ' + str(linha[dicionario["Data de emissao"][1]].year)  
                        flagCritica = False        
            
                if  dicionario[critica][8]:
                    if linha[dicionario[critica][0]] != linha[dicionario[critica][2]]:
                        if linha[dicionario[critica][0]].strip().rjust(12,'0') == linha[dicionario[critica][2]].strip().rjust(12,'0'):
                            linha[dicionario[critica][5]] = 'DIFERENTE - APENAS ZEROS A ESQUERDA'    
                            flagCritica = False 
                            dicionario[critica][9] += 1
                    
                    if linha[dicionario[critica][1]] != linha[dicionario[critica][2]]:
                        if linha[dicionario[critica][1]].strip().rjust(12,'0') == linha[dicionario[critica][2]].strip().rjust(12,'0'):
                            linha[dicionario[critica][6]] = 'DIFERENTE - APENAS ZEROS A ESQUERDA'    
                            flagCritica = False 
                            dicionario[critica][10] += 1
                
                if flagCritica:

                    if critica == "Valor_Contabil":
                        if (type(linha[dicionario[critica][0]]) == float or type(linha[dicionario[critica][0]]) == int):
                            linha[dicionario[critica][0]] = '%.2f'%float(linha[dicionario[critica][0]])
                        
                        if (type(linha[dicionario[critica][1]]) == float or type(linha[dicionario[critica][1]]) == int):
                            linha[dicionario[critica][1]] = '%.2f'%float(linha[dicionario[critica][1]])
                        
                        if (type(linha[dicionario[critica][2]]) == float or type(linha[dicionario[critica][2]]) == int):
                            linha[dicionario[critica][2]] = '%.2f'%float(linha[dicionario[critica][2]])


                    if str(linha[dicionario[critica][0]]).strip() != str(linha[dicionario[critica][2]]).strip():
                        dicionario[critica][3] += 1   
                        relatar = True
                        linha[dicionario[critica][5]] = 'DIFERENTE'       
                    else:
                        linha[dicionario[critica][5]] = 'IGUAL'
                    
                    # ORIGINAL x REGERADO

                    if str(linha[dicionario[critica][1]]).strip() != str(linha[dicionario[critica][2]]).strip():
                        dicionario[critica][4] += 1   
                        relatar = True
                        linha[dicionario[critica][6]] = 'DIFERENTE'
                    else:
                        linha[dicionario[critica][6]] = 'IGUAL'

        if dicionario["Data de emissao"]:
            x = 0
            while x < 3:
                if linha[dicionario["Data de emissao"][x]] != '':
                    data = linha[dicionario["Data de emissao"][x]].strftime('%d/%m/%Y')
                    linha[dicionario["Data de emissao"][x]] = data
                x += 1            

        if relatar:
            relatorio.write(';'.join('"'"'"+str(x)+"'"'"' for x in linha[:-1]) + '\n')
            count += 1

    
    if relatorio:
        relatorio.close()
        os.system('rm ' + nome_relatorio + '.zip')
        dir_corrente = os.getcwd()
        os.chdir(dir_relatorio)                    
        compacta = 'zip -q -n zip ' + nome_arquivo + '.zip ' + nome_arquivo  
        os.system(compacta)
        os.system('rm ' + nome_relatorio )
        os.chdir(dir_corrente)
    
    for erro in dicionario.keys():
        log('QUANTIDADE DIFERENÇA - IMPRESSÃO x REPROCESSADO - %s : %s' %( erro, dicionario[erro][3] ))
        log('QUANTIDADE DIFERENÇA - ORIGINAL  x REPROCESSADO - %s : %s' %( erro, dicionario[erro][4] ))
        if dicionario[erro][8]:
            log('QUANTIDADE DIFERENÇA APENAS ZERO A ESQUERDA - IMPRESSÃO x REPROCESSADO - %s : %s' %( erro, dicionario[erro][9] ))
            log('QUANTIDADE DIFERENÇA APENAS ZERO A ESQUERDA -  ORIGINAL x REPROCESSADO - %s : %s' %( erro, dicionario[erro][10] ))

    if relatorio:            
        log("#"*100)
        log("# ")
        log(" ARQUIVO DE SAIDA " )
        log( nome_relatorio )
    log("#"*100)

    if ret > 0:
        log("ERRO.")  

    log(" FIM DO PROCESSAMENTO...")

    return ret

def ultimodia(ano,mes):
   return(31 if mes == 12 else datetime.date.fromordinal((datetime.date(ano,mes+1,1)).toordinal()-1).day)

if __name__ == "__main__":
    cod_saida = 0
    log('-'*100)
    if len(sys.argv) < 2:
        log('#### ERRO ') 
        log('-'* 100)
        log('QUANTIDADE DE PARAMETROS INVALIDA')
        log('-'* 100)
        log('EXEMPLO')
        log('-'* 100)
        log( '%s <ID SERIE LEVANTAMENTO>'%( sys.argv[0] ) )
        log('-'* 100)
        cod_saida = 99
    else:
        log('ID_SERIE_LEVANTAMENTO',' ', str(sys.argv[1]))
        id_serie = sys.argv[1]    
        cod_saida = processar(id_serie)
    
    log('-'*100)
    if (cod_saida is None or cod_saida > 0):
        log("ERRO... Código de execução = ", cod_saida)
    sys.exit(cod_saida)
