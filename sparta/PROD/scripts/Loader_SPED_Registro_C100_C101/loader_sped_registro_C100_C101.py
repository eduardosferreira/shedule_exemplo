#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SCRIPT ......: populaC100C101.py
  CRIACAO .....: 08/02/2021
  AUTOR .......: Airton Borges da Silva Filho / KYROS TECNOLOGIA
  DESCRICAO ...: Este script le os arquivos protocolados do SPED, buscando os registros do
                 tipo C100, C101 e cadastra os valores encontrados em uma tabela no
                 banco de dados.

  DOCUMENTACAO : N/A

----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 09/02/2021 - Airton Borges da Silva Filho - Kyros Tecnologia
        - Criacao do script.
    * 12/02/2021 - Separar o parametro MMAAAA em MM e AAAA

            SCRIPT ......: loader_sped_registro_C100_C101.py
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
import atexit
import re

from pathlib import Path

import comum
import sql
import layout
import util

log.gerar_log_em_arquivo = True

status_final = 0

def processar(puf,pme,pan,pie) :
    pma=pme+pan
    comum.carregaConfiguracoes(configuracoes)
    diretorio = configuracoes.diretorio_arquivos
    tabela = configuracoes.tabela_C100_C101
    log("Gerando lista de comandos a processar ...")
    owner = 'gfcarga'
    erros = 0
    qtde = 0
    status = 0
    con=sql.geraCnxBD(configuracoes)


##########  TABELA DAS LINHAS C100

    colunas = {}
    colunas[0] = ['NOME_ARQ', 'VARCHAR2(60)', 'Primary Key']
    colunas[1] = ['IE_ENTIDADE', 'VARCHAR2(14)', 'Primary Key' ]
    colunas[2] = ['UF', 'VARCHAR2(2)', 'Primary Key' ]
    colunas[3] = ['MES_ANO', 'VARCHAR2(6)', 'Primary Key' ]
    
    colunas[4] = ['IND_OPER', 'VARCHAR2(1)', '' ]
    colunas[5] = ['IND_EMIT', 'VARCHAR2(1)', '' ]
    colunas[6] = ['COD_PART', 'VARCHAR2(60)', '' ]
    colunas[7] = ['COD_MOD', 'VARCHAR2(2)', '' ]
    colunas[8] = ['COD_SIT', 'NUMBER(2)', '' ]
    colunas[9] = ['SER', 'VARCHAR2(3)', '' ]
    colunas[10] = ['NUM_DOC', 'NUMBER(9)', '' ]
    colunas[11] = ['CHV_NFE', 'VARCHAR2(44)', '' ]
    colunas[12] = ['DT_DOC', 'NUMBER(8)', '' ]
    colunas[13] = ['DT_E_S', 'NUMBER(8)', '' ]
    colunas[14] = ['VL_DOC', 'NUMBER(19,2)', '' ]
    colunas[15] = ['IND_PGTO', 'VARCHAR2(1)', '' ]
    colunas[16] = ['VL_DESC', 'NUMBER(19,2)', '' ]
    colunas[17] = ['VL_ABAT_NT', 'NUMBER(19,2)', '' ]
    colunas[18] = ['VL_MERC', 'NUMBER(19,2)', '' ]
    colunas[19] = ['IND_FRT', 'VARCHAR2(1)', '' ]
    colunas[20] = ['VL_FRT', 'NUMBER(19,2)', '' ]
    colunas[21] = ['VL_SEG', 'NUMBER(19,2)', '' ]
    colunas[22] = ['VL_OUT_DA', 'NUMBER(19,2)', '' ]
    colunas[23] = ['VL_BC_ICMS', 'NUMBER(19,2)', '' ]
    colunas[24] = ['VL_ICMS', 'NUMBER(19,2)', '' ]
    colunas[25] = ['VL_BC_ICMS_ST', 'NUMBER(19,2)', '' ]
    colunas[26] = ['VL_ICMS_ST', 'NUMBER(19,2)', '' ]
    colunas[27] = ['VL_IPI', 'NUMBER(19,2)', '' ]
    colunas[28] = ['VL_PIS', 'NUMBER(19,2)', '' ]
    colunas[29] = ['VL_COFINS', 'NUMBER(19,2)', '' ]
    colunas[30] = ['VL_PIS_ST', 'NUMBER(19,2)', '' ]
    colunas[31] = ['VL_COFINS_ST', 'NUMBER(19,2)', '' ]
    colunas[32] = ['VL_FCP_UF_DEST', 'NUMBER(19,2)', '' ]
    colunas[33] = ['VL_ICMS_UF_DEST', 'NUMBER(19,2)', '' ]
    colunas[34] = ['VL_ICMS_UF_REM', 'NUMBER(19,2)', '' ]


    campos = {}
  
    campos['NOME_ARQ'] = 0
    campos['IE_ENTIDADE'] = 0
    campos['UF'] = 0
    campos['MES_ANO'] = 0
    
    campos['REG'] = 1
    campos['IND_OPER'] = 2
    campos['IND_EMIT'] = 3
    campos['COD_PART'] = 4
    campos['COD_MOD'] = 5
    campos['COD_SIT'] = 6
    campos['SER'] = 7
    campos['NUM_DOC'] = 8
    campos['CHV_NFE'] = 9
    campos['DT_DOC'] = 10
    campos['DT_E_S'] = 11
    campos['VL_DOC'] = 12
    campos['IND_PGTO'] = 13
    campos['VL_DESC'] = 14
    campos['VL_ABAT_NT'] = 15
    campos['VL_MERC'] = 16
    campos['IND_FRT'] = 17
    campos['VL_FRT'] = 18
    campos['VL_SEG'] = 19
    campos['VL_OUT_DA'] = 20
    campos['VL_BC_ICMS'] = 21
    campos['VL_ICMS'] = 22
    campos['VL_BC_ICMS_ST'] = 23
    campos['VL_ICMS_ST'] = 24
    campos['VL_IPI'] = 25
    campos['VL_PIS'] = 26
    campos['VL_COFINS'] = 27
    campos['VL_PIS_ST'] = 28
    campos['VL_COFINS_ST'] = 29
 
    campos['VL_FCP_UF_DEST'] = 0
    campos['VL_ICMS_UF_DEST'] = 0
    campos['VL_ICMS_UF_REM'] = 0
    
    
##########  TABELA DAS LINHAS C101  
    
    colunasC101 = {}
    colunasC101[0] = ['REG', 'VARCHAR2(4)', '']
    colunasC101[1] = ['VL_FCP_UF_DEST', 'NUMBER(19,2)', '' ]
    colunasC101[2] = ['VL_ICMS_UF_DEST', 'NUMBER(19,2)', '' ]
    colunasC101[3] = ['VL_ICMS_UF_REM', 'NUMBER(19,2)', '' ]
   
    camposC101 = {} 
    camposC101['REG'] = 1
    camposC101['VL_FCP_UF_DEST'] = 2
    camposC101['VL_ICMS_UF_DEST'] = 3
    camposC101['VL_ICMS_UF_REM'] = 4

    try :
        comando = "SELECT COUNT(1) FROM %s.%s"%( owner, tabela )
        con.executa(comando)
        res = con.fetchone()
        log("Registros existentes na tabela",tabela , " = ", res[0] )
        
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
        
        if res[0] > 0  :
            if ((puf == "*") and (pma == "**") and  (pie == "*")):
                comando = "TRUNCATE TABLE %s.%s"%( owner, tabela )
            else:
                comando = "DELETE FROM %s.%s WHERE UF = %s AND MES_ANO LIKE %s AND IE_ENTIDADE = %s"%( owner, tabela, cpuf, cpma, cpie )
            
            con.executa(comando)
            con.commit()

    except Exception as e :
        log('Criando a tabela', tabela, e )
        comando = """
CREATE TABLE %s
(
  """%( tabela )
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
        comando = "SELECT COUNT(1) FROM %s.%s"%( owner, tabela )
        con.executa(comando)
        res = con.fetchone()
        log("Registros existentes após deletar os registros a serem incluidos nesta execucao: ",tabela , " = ", res[0] )    
    except Exception as e :    
        log('Nao foi possivel contar os registros na tabela ', tabela, e )

    if (pma == "______"):
        pma = "*"


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
                                                    log("-"*100)
                                                    log("  - Varrendo sub-diretorio mes :", path_final)
                                                    mascara = "SPED_"+dir_mes+dir_ano+"_"+dir_uf+"_*_PROT*.txt"
                                                    listadeies = ies_existentes(mascara,path_final)            
                                                    for iee in listadeies:
                                                        if ((iee == pie) or (pie == "*")): 
                                                            log(" INICIO do processamento para a IE ", iee)
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
                                                                insere = False
                                                                temC101 = False
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
                                                                    if (linha.startswith('|C101|') or insere == True) :
                                                                        if (linha.startswith('|C101|')):
                                                                            valores = valores[0:-15]
                                                                            tipo_reg = linha.split('|')[1] 
                                                                            dados = linha.split('|')
                                                                            temC101 = True                                                                            
                                                                            for k in colunasC101.keys() :
                                                                                col, tipo, pk = colunasC101[k]
                                                                                if col not in ['REG'] :
                                                                                    if tipo.lower().strip()[0] == 'n' :
                                                                                        valores += ', 0.00' if not dados[camposC101[col]] else ", " + str(dados[camposC101[col]].replace(",","." ))
                                                                                    else :
                                                                                        valores += "''" if not dados[camposC101[col]] else "'%s'"%(dados[camposC101[col]].replace("'","''" ))
                                                                                else:
                                                                                    continue
                                                                            insere = True
                                                                        if (insere == True):    
                                                                            try :    
                                                                                comando = """INSERT INTO %s.%s( %s ) VALUES ( %s )"""%(owner, tabela, cols, valores )
                                                                                con.executa(comando)
                                                                                if (reg_arqs % 5000) == 0 :
                                                                                    log("      Realizando COMMIT,", reg_arqs, 'registros inseridos.')
                                                                                    con.commit()
                                                                            except Exception as e :
                                                                                log('Erro ao executar insert:', e )
                                                                                erros += 1
                                                                        insere = False
                                                                    if (linha.startswith('|C100|')) :
                                                                        tipo_reg = linha.split('|')[1] 
                                                                        qtde += 1
                                                                        reg_arqs += 1
                                                                        dados = linha.split('|')
                                                                        cols = ""
                                                                        valores = ""
                                                                        for k in colunas.keys():
                                                                            col, tipo, pk = colunas[k]
                                                                            if cols :
                                                                                cols += ', '
                                                                                valores += ', '
                                                                            cols += col
                                                                            if col not in ['REG',
                                                                                           'NOME_ARQ',
                                                                                           'IE_ENTIDADE' , 
                                                                                           'UF', 'MES_ANO' , 
                                                                                           'VL_FCP_UF_DEST', 
                                                                                           'VL_ICMS_UF_DEST',
                                                                                           'VL_ICMS_UF_REM'] :                                                 
                                                                                if tipo.lower().strip()[0] == 'n' :
                                                                                    valores += '0' if not dados[campos[col]] else str(dados[campos[col]].replace(",","." ))
                                                                                else :
                                                                                    valores += "''" if not dados[campos[col]] else "'%s'"%(dados[campos[col]].replace("'","''" ))
                                                                            else :
                                                                                if col == 'NOME_ARQ' :
                                                                                    valores += "'%s'"%( item )
                                                                                elif col == 'IE_ENTIDADE' :
                                                                                    valores += "'%s'"%( ie )
                                                                                elif col == 'UF' :
                                                                                    valores += "'%s'"%( uf )
                                                                                elif col == 'MES_ANO' :
                                                                                    valores += "'%s'"%( mes_ano )
                                                                                elif col == 'VL_FCP_UF_DEST' :
                                                                                    valores += "'%s'"%( 0 )
                                                                                elif col == 'VL_ICMS_UF_DEST' :
                                                                                    valores += "'%s'"%( 0 )
                                                                                elif col == 'VL_ICMS_UF_REM' :
                                                                                    valores += "'%s'"%( 0 )
                                                                                    
                                                                        insere = True
                                                                if (insere == True):    
                                                                    try :    
                                                                        comando = """INSERT INTO %s.%s( %s ) VALUES ( %s )"""%(owner, tabela, cols, valores )
                                                                        con.executa(comando)
                                                                        if (reg_arqs % 10000) == 0 :
                                                                            log("      Realizando COMMIT,", reg_arqs, 'registros inseridos.')
                                                                            con.commit()
                                                                    except Exception as e :
                                                                        log('Erro ao executar insert C100 :', e )
                                                                        erros += 1
                                                                    insere = False
                                                                fd.close()
                                                                log("      Realizando COMMIT, totalizando", reg_arqs, 'registros C100 inseridos para esse arquivo.')
                                                                con.commit()
                                                                log(" ")
                                                                log(" ")
    
    try :
        comando = "SELECT COUNT(1) FROM %s.%s"%( owner, tabela )
        con.executa( comando )
        res = con.fetchone()
        log("Resultado final: Nome da tabela, quantidade de registros = ",tabela , ", ", res[0] )    
    except Exception as e :    
        log('Nao foi possivel contar os registros na tabela ', tabela, e )
    con.commit()

    log("="*80)
    log("Foram realizadas %s insersoes na tabela %s.%s ."%(qtde, owner, tabela))
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
        log('#### Outros exemplos validos:')  
        log('####      %s SP "" "" ""'%(sys.argv[0]))         
        log('####      %s SP 06 2020 ""'%(sys.argv[0]))         
        log('####      %s SP "" 2020 ""'%(sys.argv[0]))         
        log('####      %s "" "" 2020 ""'%(sys.argv[0]))         
        log("#### ")
        log('#### ')
        log("-" * 100)
        log("")
        return(False,False,False,False)

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
        print("parametros invalidos")
        status_final = 'ERRO'

    if not status_final :
        status_final = processar(ufi,mesi,anoi,iei)

    
