#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SCRIPT ......: populaC170.py
  CRIACAO .....: 08/03/2021
  AUTOR .......: AIRTON BORGES DA SILVA FILHO / KYROS TECNOLOGIA
  DESCRICAO ...: Este script le os arquivos protocolados do SPED de acordo com os parametros 
                  passados, buscando os registros do tipo C170 e popula uma tabela no 
                  banco de dados definida no arquivo populaC170.cfg

  DOCUMENTACAO : N/A

----------------------------------------------------------------------------------------------
  HISTORICO : 
            SCRIPT ......: loader_sped_registro_C170.py
            AUTOR .......: Victor Santos
            Alteração para novo formato de script
----------------------------------------------------------------------------------------------
"""

import os
import sys

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)

import configuracoes

import datetime
import re
from pathlib import Path

import comum
import sql
import layout
import util

status_final = 0

log.gerar_log_em_arquivo = True

def processar(puf,pme,pan,pie) :
    pma=pme+pan
    comum.carregaConfiguracoes(configuracoes)
    diretorio = configuracoes.diretorio_arquivos
    tabela_temporaria = configuracoes.tabela_dados
    log("Gerando lista de comandos a processar ...")
    owner = 'gfcarga'
    con=sql.geraCnxBD(configuracoes)
    
    erros = 0
    qtde = 0
    status = 0
 
    colunas      =  {}
    colunas[0]   =  ['NOME_ARQ'     , 'varchar2(60)' , 'Primary Key']
    colunas[1]   =  ['IE_ENTIDADE'  , 'varchar2(14)' , 'Primary Key' ]
    colunas[2]   =  ['UF'           , 'varchar2(2)'  , 'Primary Key' ]
    colunas[3]   =  ['MES_ANO'      , 'varchar2(6)'  , 'Primary Key' ]
                   
    colunas[4]   =  ['NUM_ITEM'         , 'integer'      , '' ]
    colunas[5]   =  ['COD_ITEM'         , 'varchar2(60)' , '' ]
    colunas[6]   =  ['DESCR_COMPL'      , 'varchar2(200)', '' ]
    colunas[7]   =  ['QTD'              , 'number(19,5)', '' ]
    colunas[8]   =  ['UNID'             , 'varchar2(6)', '' ]
    colunas[9]   =  ['VL_ITEM'          , 'number(19,2)', '' ]
    colunas[10]  =  ['VL_DESC'          , 'number(19,2)', '' ]
    colunas[11]  =  ['IND_MOV'          , 'varchar2(1)', '' ]
    colunas[12]  =  ['CST_ICMS'         , 'integer', '' ]
    colunas[13]  =  ['CFOP'             , 'integer', '' ]
    colunas[14]  =  ['COD_NAT'          , 'varchar2(10)', '' ]
    colunas[15]  =  ['VL_BC_ICMS'       , 'number(19,2)', '' ]
    colunas[16]  =  ['ALIQ_ICMS'        , 'number(6,2)', '' ]
    colunas[17]  =  ['VL_ICMS'          , 'number(19,2)', '' ]
    colunas[18]  =  ['VL_BC_ICMS_ST'    , 'number(19,2)', '' ]
    colunas[19]  =  ['ALIQ_ST'          , 'number(6,2)', '' ]
    colunas[20]  =  ['VL_ICMS_ST'       , 'number(19,2)', '' ]
    colunas[21]  =  ['IND_APUR'         , 'varchar2(1)', '' ]
    colunas[22]  =  ['CST_IPI'          , 'varchar2(2)', '' ]
    colunas[23]  =  ['COD_ENQ'          , 'varchar2(3)', '' ]
    colunas[24]  =  ['VL_BC_IPI'        , 'number(19,2)', '' ]
    colunas[25]  =  ['ALIQ_IPI'         , 'number(6,2)', '' ]
    colunas[26]  =  ['VL_IPI'           , 'number(19,2)', '' ]
    colunas[27]  =  ['CST_PIS'          , 'integer', '' ]
    colunas[28]  =  ['VL_BC_PIS'        , 'number(19,2)', '' ]
    colunas[29]  =  ['ALIQ_PIS'         , 'number(6,2)', '' ]
    colunas[30]  =  ['QUANT_BC_PIS'     , 'number(19,2)', '' ]
    colunas[31]  =  ['ALIQ_PIS_R'         , 'number(8,4)' , '' ]
    colunas[32]  =  ['VL_PIS'           , 'number(19,2)'  , '' ]
    colunas[33]  =  ['CST_COFINS'       , 'integer'      , '' ]
    colunas[34]  =  ['VL_BC_COFINS'     , 'number(19,2)'  , '' ]
    colunas[35]  =  ['ALIQ_COFINS'      , 'number(8,4)'  , '' ]
    colunas[36]  =  ['QUANT_BC_COFINS'  , 'number(19,2)'      , '' ]
    colunas[37]  =  ['ALIQ_COFINS_R'      , 'number(6,2)'  , '' ]
    colunas[38]  =  ['VL_COFINS'        , 'number(19,2)' , '' ]
    colunas[39]  =  ['COD_CTA'          , 'varchar2(60)'    , '' ]
    colunas[40]  =  ['VL_ABAT_NT'       , 'number(19,2)'    , '' ]

    
    campos                 =  {}
     
    campos['NOME_ARQ']          =  0
    campos['IE_ENTIDADE']       =  0
    campos['UF']                =  0
    campos['MES_ANO']           =  0
    

    campos['NUM_ITEM']          =  2
    campos['COD_ITEM']          =  3 
    campos['DESCR_COMPL']       =  4 
    campos['QTD']               =  5  
    campos['UNID']              =  6 
    campos['VL_ITEM']           =  7 
    campos['VL_DESC']           =  8 
    campos['IND_MOV']           =  9
    campos['CST_ICMS']          =  10 
    campos['CFOP']              =  11
    campos['COD_NAT']           =  12 
    campos['VL_BC_ICMS']        =  13
    campos['ALIQ_ICMS']         =  14
    campos['VL_ICMS']           =  15
    campos['VL_BC_ICMS_ST']     =  16
    campos['ALIQ_ST']           =  17
    campos['VL_ICMS_ST']        =  18
    campos['IND_APUR']          =  19
    campos['CST_IPI']           =  20
    campos['COD_ENQ']           =  21
    campos['VL_BC_IPI']         =  22
    campos['ALIQ_IPI']          =  23
    campos['VL_IPI']            =  24
    campos['CST_PIS']           =  25
    campos['VL_BC_PIS']         =  26
    campos['ALIQ_PIS']          =  27
    campos['QUANT_BC_PIS']      =  28
    campos['ALIQ_PIS_R']        =  29
    campos['VL_PIS']            =  30
    campos['CST_COFINS']        =  31
    campos['VL_BC_COFINS']      =  32
    campos['ALIQ_COFINS']       =  33
    campos['QUANT_BC_COFINS']   =  34
    campos['ALIQ_COFINS_R']     =  35
    campos['VL_COFINS']         =  36
    campos['COD_CTA']           =  37
    campos['VL_ABAT_NT']        =  38

    
    log("Gerando lista de arquivos a processar ...")

    try :
        comando = "SELECT COUNT(1) FROM %s.%s"%( owner, tabela_temporaria )
        con.executa(comando)
        res = con.fetchone()

        log("Registros existentes na tabela",tabela_temporaria , " = ", res[0] )

        if (puf == "*"):
            cpuf = "UF"
        else:
            cpuf = "'" + puf + "'"
        
        cpma = "'" + pma + "'"
        
        if (pme == "__"):
            pme = "*"
        if (pan == "____"):
            pan = "*"
           
        if (pie == "*"):
            cpie = "IE_ENTIDADE"
        else:
            cpie = "'" + pie + "'"
    
        if res[0] > 0 :
            if ((puf == "*") and (pma == "**") and  (pie == "*")):
                comando = "TRUNCATE TABLE %s.%s"%( owner, tabela_temporaria )
            else:
                comando = "DELETE FROM %s.%s WHERE UF = %s AND MES_ANO LIKE %s AND IE_ENTIDADE = %s"%( owner, tabela_temporaria, cpuf, cpma, cpie )
            con.executa(comando)
            con.commit()

    except Exception as e :
        log('Criando tabela', tabela_temporaria, e )
        comando = """
CREATE TABLE %s
(
  """%( tabela_temporaria )
        txt = """"""
        primaryKeys = ""
        for k in colunas.keys() :
            col, tipo, pk = colunas[k]
            if txt :
                txt += ',\n' 
            txt += col + '  ' + tipo
        comando += txt
        comando += """ --,
    --PRIMARY KEY ( %s )
)"""%( primaryKeys )

        con.executa(comando)
        con.commit()
    try :
        comando = "SELECT COUNT(1) FROM %s.%s"%( owner, tabela_temporaria )
        con.executa(comando)
        res = con.fetchone()
        log("Registros existentes após deletar os registros a serem incluidos nesta execucao: ",tabela_temporaria , " = ", res[0] )    
    except Exception as e :    
        log('Nao foi possivel contar os registros na tabela ', tabela_temporaria, e )        
        

    if (pma == "______"):
        pma = "*"
    
    if not os.path.isdir(diretorio) :
        os.makedirs(diretorio) 
 
    for dir_uf in os.listdir( diretorio ) :
        if util.validauf(dir_uf):
            if ((dir_uf == puf) or (puf == "*")):
                path = os.path.join(diretorio, dir_uf)
                if os.path.isdir(path) :
                    log("#"* 100)
                    log("="* 100)                
                    log("Varrendo diretorio :", path)
                    for dir_ano in os.listdir( path ) :
                        if util.validaano(dir_ano):
                            if ((dir_ano == pma[2:]) or (pma == "*") or (pan == "*") or pan == "____" or pma == "______"):                                
                                path_ano = os.path.join(path, dir_ano)
                                if os.path.isdir(path_ano) :
                                    log("  - Varrendo sub-diretorio ano :", path_ano)
                                    
                                    for dir_mes in os.listdir( path_ano ) :
                                        if util.validames(dir_mes):
                                            if ((dir_mes == pma[0:2]) or (pme == "*") or (pme == "__") or pma == "______"):
                                                path_final = os.path.join(path_ano, dir_mes)
                                                if os.path.isdir(path_final) :
                                                    log(" ")
                                                    log("-"*100)
                                                    log("  - Varrendo sub-diretorio mes :", path_final)
                        
                                                    mascara = "SPED_"+dir_mes+dir_ano+"_"+dir_uf+"_*_PROT*.txt"
                                                    listadeies = ies_existentes(mascara,path_final)
                                                    for iee in listadeies:
                                                        if ((iee == pie) or (pie == "*")): 
                                                            log(" INÍCIO do processamento para a IE ", iee)
                                                            mascara_protocolado = "SPED_"+dir_mes+dir_ano+"_"+dir_uf+"_"+iee+"_PROT*.txt"
                                                            nome_protocolado = util.nome_arquivo(mascara_protocolado,path_final)
                                                            item = str(nome_protocolado).split('/')[-1] 
                                                            path_arquivo = nome_protocolado
                                                            if os.path.isfile(path_arquivo) :
                                                                log('     Lendo arquivo :', item)
                                                                reg_arqs = 0
                                                                fd = open(path_arquivo, 'r', encoding=comum.encodingDoArquivo(path_arquivo))
                                                                uf = None
                                                                mes_ano = None
                                                                for linha in fd :
                                                                    if linha.startswith('|0000|') :
                                                                        uf = linha.split('|')[9]
                                                                        mes_ano = linha.split('|')[4][-6:]
                                                                        ie = linha.split('|')[10]
                                                                        log("     - Referente a UF ...:", uf)
                                                                        log("     - Referente ao mes .:", mes_ano)
                                                                        log("     - Referente a IE ...:", ie)
                                                                        if uf != dir_uf :
                                                                            log('ERRO : UF do registro (0000) header, divergente do diretorio do arquivo.')
                                                                            status += 1
                                                                            break
                                                                        if not item.__contains__(uf) :
                                                                            log('ERRO : UF do registro (0000) header, divergente da nomenclatura do arquivo.')
                                                                            status += 1
                                                                            break
                                                                        if not item.__contains__(mes_ano) :
                                                                            log('ERRO : Data de referencia do registro (0000) header, divergente da nomenclatura do arquivo.')
                                                                            status += 1
                                                                            break
                                                                        if not item.__contains__(ie) :
                                                                            log('ERRO : IE do registro (0000) header, divergente da nomenclatura do arquivo.')
                                                                            status += 1
                                                                            break
                            
                                                                    elif linha.startswith('|C170|') :
                                                                        qtde += 1
                                                                        reg_arqs += 1
                                                                        try :
                                                                            dados = linha.split('|')
                                                                            cols = ""
                                                                            valores = ""
                                                                            for k in colunas.keys() :
                                                                                col, tipo, pk = colunas[k]
                                                                                if cols :
                                                                                    cols += ', '
                                                                                    valores += ', '
                                                                                cols += col
                                                                                if col not in [ 'NOME_ARQ', 'IE_ENTIDADE', 'UF', 'MES_ANO', 'REG'  ] :
                                                                                    if (dados[campos[col]] == '\n'):
                                                                                        dados[campos[col]]='0'
                                                                                    if (not tipo.lower().strip()[0] == 'v'):
                                                                                        valores += '0' if not dados[campos[col]] else str(dados[campos[col]].replace(",","." ))
                                                                                    else :
                                                                                        valores += "''" if not dados[campos[col]] else "'%s'"%(dados[campos[col]].replace("'","''" ).replace(",","." ))
                                                                                else :
                                                                                    if col == 'NOME_ARQ' :
                                                                                        valores += "'%s'"%( item )
                                                                                    elif col == 'IE_ENTIDADE' :
                                                                                        valores += "'%s'"%( iee )
                                                                                    elif col == 'UF' :
                                                                                        valores += "'%s'"%( uf )
                                                                                    elif col == 'MES_ANO' :
                                                                                        valores += "'%s'"%( mes_ano )
                                                                                    elif col == 'REG' :
                                                                                        continue
                                                                            comando = """INSERT INTO %s.%s ( %s ) VALUES ( %s )"""%(owner, tabela_temporaria, cols, valores )
                                                                            con.executa(comando)
                                                                            if (reg_arqs % 10000) == 0 :
                                                                                log("      Realizando COMMIT,", reg_arqs, 'registros inseridos.')
                                                                                con.commit()
                                                                        except Exception as e :
                                                                            log('Erro ao executar insert :', e )
                                                                            erros += 1
                                                                fd.close()
                                                                log("      Realizando COMMIT, totalizando", reg_arqs, 'registros inseridos para esse arquivo.')
                                                                con.commit()
                
    try :
        comando = "SELECT COUNT(1) FROM %s.%s"%( owner, tabela_temporaria )
        con.executa(comando)
        res = con.fetchone()
        log("Resultado final: Nome da tabela, quantidade de registros = ",tabela_temporaria , ", ", res[0] )    
    except Exception as e :    
        log('Nao foi possivel contar os registros na tabela ', tabela_temporaria, e )   
    
    con.commit()    

    log("="*80)
    log("Foram realizadas %s inserções na tabela %s.%s ."%(qtde, owner, tabela_temporaria))
    log(" - %s com erros."%(erros))
    log(" - %s com sucesso."%(qtde - erros))
    log(" - %s erros relativos aos arquivos processados."%(status))
    log("="*80)

    return erros + status

def recebeparametros():
#### Recebe, verifica e formata os argumentos de entrada.
    ret = 0
    ufi = ""
    mesanoi = ""
    mesi = ""
    anoi = "" 
    iei = ""
    if (len(sys.argv) == 5):
        if (len(sys.argv[1]) >0 ): ufi = util.valida_uf(str(sys.argv[1]))
        else: ufi = "*"
        if (len(sys.argv[2]) > 0): mesi = util.valida_mes(str(sys.argv[2])) 
        else: mesi = "__"
        if (len(sys.argv[3]) > 0): anoi = util.valida_ano(str(sys.argv[3])) 
        else: anoi = "____"
        if (len(sys.argv[4]) > 0): iei = util.valida_ie(str(sys.argv[4]))
        else: iei = "*"
    else:
        ret = 99
    if ( ufi == "#" or mesi == "#" or anoi == "#"or iei == "#" ):
        ret = 99
    if ( ret != 0):
        log("-" * 100)
        log("#### ")
        log('#### ERRO - Erro nos parametros do script.')
        log("#### ")
        log('#### Exemplo de como deve ser :')
        log('####      %s [UF] [MM] [AAAA] [IE] '%(sys.argv[0]))
        log("#### ")
        log('#### Onde')
        log('####      [UF]     = Opcional. Estado. Ex: SP, MG, RJ, PE, [OUTRO ESTADO]  ou para todos, informe "".')
        log('####      [MM]     = Opcional. Mes. Ex: Para junho, informe 06 ou para todos, informe "". ')
        log('####      [AAAA]   = Opcional. Ano. Ex: Para 2020 informe 2020 ou para todos, informe "". ')
        log('####      [IE]     = Opcional. Inscricao Estadual. Ex: 108383949112 ou "".')
        log("#### ")
        log('#### Portanto, se o estado = SP, o mes = 06 e o ano = 2020, e IE = 108383949112 o comando correto deve ser :')  
        log('####      %s SP 06 2020 108383949112'%(sys.argv[0]))  
        log('#### ')
        log('#### Outros exemplos validos:')  
        log('####      %s SP "" "" ""   - Todos os registros de SP serão apagados e carregados novamente dos arquivos disponíveis.'%(sys.argv[0]))         
        log('####      %s SP 06 2020 "" - Os registros de SP, mes 06 de 2020 serão apagados e carregados novamente do arquivo disponível.'%(sys.argv[0]))         
        log('####      %s SP "" 2020 "" - Os registros de SP, do ano 2020, de todos os meses, serão apagados e carregados novamente dos arquivos disponíveis.'%(sys.argv[0]))         
        log('####      %s "" "" 2020 "" - Os registros de todas UFs, todos os meses de 2020 serão apagados e carregados novamente dos arquivos disponíveis.'%(sys.argv[0])) 
        log('####      %s "" "" "" ""   - Toda a tabela será apagada e a carga será refeita com todas UFs, MESES e ANOS disponíveis.'%(sys.argv[0]))       
        log("#### ")

        log("-" * 100)
        log("")
        return(False,False,False,False)

    print (ufi,mesi,anoi,iei)
    return (ufi,mesi,anoi,iei)

def ies_existentes(mascara,diretorio):
    global ret
    
    qdade = 0
    ies = []
    directory = Path(diretorio)
    files = directory.glob(mascara)
    sorted_files = sorted(files, reverse=False)
    if sorted_files:
        log("Arquivos encontrados: ")
        for f in sorted_files:
            qdade = qdade + 1
            ie = os.path.basename(str(f)).split("_")[3]
            log("   ",qdade, " => ", f)
            try:
                ies.index(ie)
            except:
                ies.append(ie)
                continue
            
    else: 
        log('ERRO : Arquivo %s não está na pasta %s'%(mascara,diretorio))
        log("")
        ret=99
        return("")
    log(" ")
    return(ies)

def inicializar() :
    ret = 0
    dic_layouts = layout.carregaLayout()
    if not ret and not dic_layouts :
        ret = 2
    configuracoes.dic_layouts = dic_layouts
    if not configuracoes.banco:
        log("Erro falta variavel 'banco' no arquivo de configuração (.cfg).")
        ret = 1
    return ret

if __name__ == "__main__":
    status_final = 0
    ufi,mesi,anoi,iei = recebeparametros()
    if ufi and mesi and anoi and iei:
        status_final = inicializar()
    else:
        log("ERRO - parametros invalidos")
        status_final = 'ERRO'

    if not status_final :
        status_final = processar(ufi,mesi,anoi,iei)