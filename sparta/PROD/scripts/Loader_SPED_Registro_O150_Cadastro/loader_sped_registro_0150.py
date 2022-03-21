#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SCRIPT ......: deParaCadastroSPED.py
  CRIACAO .....: 02/02/2021
  AUTOR .......: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO ...: Este script le os arquivos protocolados do SPED, buscando os registros do
                 tipo 0150, e cadastra os valores encontrados em uma tabela temporaria no
                 banco de dados.

  DOCUMENTACAO : N/A

----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 02/02/2021 - Welber Pena de Sousa - Kyros Tecnologia
        - Criacao do script.
        
    * 09/02/2021 - Airton Borges da Silva Filho - Kyros Tecnologia
        - Alterado para inserir no banco somente os registros dos arquivos válidos

    * 18/08/2021 - Adequação para novo formato de script 
        SCRIPT ......: loader_sped_registro_O150.py
        AUTOR .......: Victor Santos
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

name_script = os.path.basename(__file__).split('.')[0].split('_')[0]

status_final = 0

log.gerar_log_em_arquivo = True

SD = ('/' if os.name == 'posix' else '\\')

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

    colunas = {}
    colunas[0] = ['NOME_ARQ', 'varchar2(60)', 'Primary Key']
    colunas[1] = ['IE_ENTIDADE', 'varchar2(14)', 'Primary Key' ]
    colunas[2] = ['UF', 'varchar2(2)', 'Primary Key' ]
    colunas[3] = ['MES_ANO', 'varchar2(6)', 'Primary Key' ]
   
    colunas[4] = ['COD_PART', 'varchar2(60)', '' ]
    colunas[5] = ['NOME', 'varchar2(100)', '' ]
    colunas[6] = ['COD_PAIS', 'integer', '' ]
    colunas[7] = ['CNPJ', 'integer', '' ]
    colunas[8] = ['CPF', 'integer', '' ]
    colunas[9] = ['IE_CLI_FOR', 'varchar2(14)', '' ]
    colunas[10] = ['COD_MUN', 'integer', '' ]
    colunas[11] = ['SUFRAMA', 'varchar2(9)', '' ]
    colunas[12] = ['END', 'varchar2(60)', '' ]
    colunas[13] = ['NUM', 'varchar2(10)', '' ]
    colunas[14] = ['COMPL', 'varchar2(60)', '' ]
    colunas[15] = ['BAIRRO', 'varchar2(60)', '' ]

    campos = {}
    campos['UF'] = 0 
    campos['MES_ANO'] = 1
    campos['COD_PART'] = 2
    campos['NOME'] = 3
    campos['COD_PAIS'] = 4
    campos['CNPJ'] = 5
    campos['CPF'] = 6
    campos['IE_CLI_FOR'] = 7
    campos['COD_MUN'] = 8
    campos['SUFRAMA'] = 9
    campos['END'] = 10
    campos['NUM'] = 11
    campos['COMPL'] = 12
    campos['BAIRRO'] = 13


    log("Gerando lista de arquivos a processar ...")

    try :
        comando = "SELECT COUNT(1) FROM %s.%s"%( owner,tabela_temporaria )
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
                comando = "TRUNCATE TABLE %s.%s"%( owner,tabela_temporaria )
            else:
                comando = "DELETE FROM %s.%s WHERE UF = %s AND MES_ANO LIKE %s AND IE_ENTIDADE = %s"%( owner,tabela_temporaria, cpuf, cpma, cpie )
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
        comando = "SELECT COUNT(1) FROM %s.%s"%( owner,tabela_temporaria )
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
                            
                                                                    elif linha.startswith('|0150|') :
                                                                        qtde += 1
                                                                        reg_arqs += 1
                                                                        try :
                                                                            dados = linha.split('|')
                                                                            cols = ""
                                                                            valores = ""
                                                                            dados[0] = uf 
                                                                            dados[1] = mes_ano
                                                                            for k in colunas.keys() :
                                                                                col, tipo, pk = colunas[k]
                                                                                if cols :
                                                                                    cols += ', '
                                                                                    valores += ', '
                                                                                cols += col
                                                                                if col not in [ 'NOME_ARQ', 'IE_ENTIDADE' ] :
                                                                                    if tipo.lower().strip() == 'integer' :
                                                                                        valores += '0' if not dados[campos[col]] else str(dados[campos[col]])
                                                                                    else :
                                                                                        valores += "''" if not dados[campos[col]] else "'%s'"%(dados[campos[col]].replace("'","''" ))
                                                                                else :
                                                                                    if col == 'NOME_ARQ' :
                                                                                        valores += "'%s'"%( item )
                                                                                    elif col == 'IE_ENTIDADE' :
                                                                                        valores += "'%s'"%( ie )
                            
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
        comando = "SELECT COUNT(1) FROM %s.%s"%( owner,tabela_temporaria )
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


