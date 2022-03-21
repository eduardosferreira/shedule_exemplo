#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: cargaProtocolado.py
  CRIACAO ..: 08/02/2021
  AUTOR ....: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO : Processo para realizar a carga dos arquivos de Mestre, Itens e Destinatario .
                Parametros :
                    - UF 
                    - Periodo
                    - Serie ( Opcional )
                a partir da UF, período e Série (opcional, caso não informado considera todas) 
                informado identifica o diretório protocolado (usando tabela tsh_serie_levantamento) 
                e realiza a execução das linhas de comando para o respectivo CTL (loader - sqlldr)

----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 08/02/2021 - Welber Pena de Sousa - Kyros Tecnologia
        - Criacao do script.

    * 24/08/2021 - Adequação para novo formato de script 
        SCRIPT ......: loader_sped_registro_O150.py
        AUTOR .......: Victor Santos
    
    * 20/01/2022 - #### ALT001 
        Fazer com que o script carregue os arquivos de CONTROLE
        AUTOR .......: Welber Pena
    
----------------------------------------------------------------------------------------------
"""

import os
import sys

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
dir_execucao = os.path.realpath('.')
sys.path.append(dir_base)
import configuracoes

import time
import cx_Oracle
import datetime
import calendar


import threading
import traceback

import comum
import sql
import layout
import util

log.gerar_log_em_arquivo = True

idControle = ''
codigoUF = ''
LayoutDestinatario = ''
LayoutMestre   = ''
dic_registros = {}
dic_layouts = {}
dic_campos = {}
variaveis = {}
dic_fd = {}
listaMestreProtocolado = []
#### ALT001 - Inicio - Adequar script para carregar CONTROLES
listaDestinatarioProtocolado = []
listaControleProtocolado = []
#### ALT001 - Fim

def processar():     
            
    global listaMestreProtocolado
    #### ALT001 - Inicio - Adequar script para carregar CONTROLES
    global listaDestinatarioProtocolado 
    global listaControleProtocolado
    #### ALT001 - Fim

    comum.carregaConfiguracoes(configuracoes)

    log("Iniciando carga do(s) destinatario(s) a partir do(s) arquivo(s) Protocolado(s) ...")    
    
    dt_ref = variaveis['DT_REF'] 
    uf = variaveis['UF']
    serie = ''
    if variaveis['Serie'] :
        serie = variaveis['Serie'].replace(' ', '').replace(',', "','")

    dt_ref_inicio = "01/%s"%(dt_ref)
    dt_ref_inicio = dt_ref_inicio.replace('/','')

    dia_semana, ult_dia = calendar.monthrange(int(dt_ref.split('/')[1]), int(dt_ref.split('/')[0]) )
    dt_ref_fim = "%s/%s"%( ult_dia, dt_ref )
    dt_ref_fim = dt_ref_fim.replace('/','')

    log('- Buscando diretorios de protocolados a carregar .')
    comando = """
        SELECT TO_CHAR(l.MES_ANO, 'MM')    MES
            , TO_CHAR(l.MES_ANO, 'YYYY')   ANO_AAAA
            , SERIE                        SERIE_ORIGINAL
            , REPLACE(SERIE,' ','')        SERIE
            , l.FILI_COD                   FILIAL
            , f.UNFE_SIG                   UF
            , l.ID_SERIE_LEVANTAMENTO      ID_SERIE
            , '/portaloptrib/LEVCV115/' || f.UNFE_SIG||'/'||TO_CHAR(l.MES_ANO, 'YY/MM')||'/'||l.EMPS_COD||'/'||l.FILI_COD||'/SERIE/'||l.ID_SERIE_LEVANTAMENTO DIRETORIO
        FROM gfcarga.TSH_SERIE_LEVANTAMENTO l
            INNER JOIN OPENRISOW.FILIAL f ON l.EMPS_COD = f.EMPS_COD AND l.FILI_COD = f.FILI_COD
        WHERE MES_ANO >= to_date('%s', 'ddmmyyyy' )
            AND MES_ANO <= to_date('%s', 'ddmmyyyy' )
            AND UNFE_SIG ='%s'
"""%( dt_ref_inicio, dt_ref_fim, uf)    

    if serie :
        comando += """            AND REPLACE(SERIE,' ','')    IN ('%s')
"""%( serie )
    comando += """        order by TO_CHAR(MES_ANO, 'YYYY'), TO_CHAR(MES_ANO, 'MM'), SERIE
    """
    print(comando)

    con=sql.geraCnxBD(configuracoes)

    log('Iniciando processamento ...')
    registros_processados             = 0

    try :
        con.executa(comando)
        records = con.fetchall() 
        
    except Exception as e :
        log('Erro ao executar o SELECT abaixo')
        log(comando)
        return 3

    cols_sql = {}
    colunas = con.description()
    for x in range(len(colunas)) :
        nome = colunas[x][0]
        cols_sql[nome] = x

    while records :
        row = records.pop(0)
        registros_processados += 1
        
        ano_aaaa         = row[cols_sql['ANO_AAAA']]
        mes              = row[cols_sql['MES']]
        serie            = row[cols_sql['SERIE']]   
        filial           = row[cols_sql['FILIAL']]   
        serie_original   = row[cols_sql['SERIE_ORIGINAL']]   
        uf               = row[cols_sql['UF']]   
        id_serie         = row[cols_sql['ID_SERIE']]
        dir_serie        = row[cols_sql['DIRETORIO']]   
        dir_protocolado  = dir_serie + '/ULTIMA_ENTREGA'

        if not os.path.isdir(dir_protocolado) :
            log('Diretorio da serie %s NAO EXISTE ... Ignorando .... '%(serie))
            log('Diretorio ..: %s'%(dir_protocolado))
            continue
        lst_dir = os.listdir(dir_protocolado)
        log("# Carregando arquivos da Serie %s de %s/%s"%(serie, mes, ano_aaaa))
        log("- diretorio de arquivos :", dir_protocolado)
        list_arqs_mestre = []
        list_arqs_itens = []
        
        #### ALT001 - Inicio - Adequar script para carregar CONTROLES
        list_arqs_destinatario = []
        list_arqs_controle = []
        #### ALT001 - Fim

        pos = 28 if int(ano_aaaa) >= 2017 else 10

        lst_threads = []
        for arq in lst_dir :
            if len(arq) > pos and os.path.isfile( os.path.join( dir_protocolado, arq ) ) : 
                # print(arq[pos])
                if arq[pos] == 'M' :
                    list_arqs_mestre.append(arq)
                elif arq[pos] == 'I' :
                    list_arqs_itens.append(arq)
                #### ALT001 - Inicio - Adequar script para carregar CONTROLES
                elif arq[pos] == 'D' :
                    list_arqs_destinatario.append(arq)
                elif arq[pos] == 'C' :
                    list_arqs_controle.append(arq)
                #### ALT001 - Fim
        
        #### ALT001 - Inicio - Adequar script para carregar CONTROLES
        if len(list_arqs_mestre) != len(list_arqs_itens) or len(list_arqs_mestre) != len(list_arqs_destinatario) or len(list_arqs_mestre) != len(list_arqs_controle) :
            log('Erro : Quantide de volumes divergentes ... verifique !')
            continue
        if len(list_arqs_itens) == 0 :
            log('Erro : Falta arquivos de itens ... verifique !')
            continue
        if len(list_arqs_mestre) == 0 :
            log('Erro : Falta arquivos mestre ... verifique !')
            continue
        #### ALT001 - Inicio - Adequar script para carregar CONTROLES
        if len(list_arqs_controle) == 0 :
            log('Erro : Falta arquivos controle ... verifique !')
            continue
        #### ALT001 - Fim

        if variaveis['Processar'] in [ 'TODOS', 'MESTRE' ] :
            th = threading.Thread( target = cargaProtocolado, args = ( dir_protocolado, list_arqs_mestre, ano_aaaa, mes, serie, uf, id_serie, os.getpid() ) )
            th.start()
            lst_threads.append(th)
            time.sleep(2)

        #### ALT001 - Inicio - Adequar script para carregar CONTROLES
        if variaveis['Processar'] in [ 'TODOS', 'DESTINATARIO' ] :
            th = threading.Thread( target = cargaProtocolado, args = ( dir_protocolado, list_arqs_destinatario, ano_aaaa, mes, serie, uf, id_serie, os.getpid() ) )
            #### ALT001 - Fim
            th.start()
            lst_threads.append(th)
            time.sleep(2)
        
        if variaveis['Processar'] in [ 'TODOS', 'ITEM' ] :
            th = threading.Thread( target = cargaProtocolado, args = ( dir_protocolado, list_arqs_itens, ano_aaaa, mes, serie, uf, id_serie, os.getpid() ) )
            th.start()
            lst_threads.append(th)
            time.sleep(2)
        
        #### ALT001 - Inicio - Adequar script para carregar CONTROLES
        if variaveis['Processar'] in [ 'TODOS', 'CONTROLE' ] :
            th = threading.Thread( target = cargaProtocolado, args = ( dir_protocolado, list_arqs_controle, ano_aaaa, mes, serie, uf, id_serie, os.getpid() ) )
            th.start()
            lst_threads.append(th)
            time.sleep(2)
        #### ALT001 - Fim

        while threading.active_count() > 1 :
            time.sleep(10)
        
    log("-------------------------------------------------------------------------")
    log('*', " R E S U M O ".center(69,'-') , '*')
    log("* registros_processados ...............:", str(registros_processados))
    log("-------------------------------------------------------------------------")

    log("* Processamento Finalizado !")
    log("-------------------------------------------------------------------------")
    retorno = 0 if not os.path.isfile( 'ERRO_%s'%(os.getpid())) else 5
    if retorno :
        os.remove('ERRO_%s'%(os.getpid()))

    return retorno


def cargaProtocolado(diretorio, lst_arquivos, ano, mes, serie, uf, id_serie, pid ):

    os.chdir( log.path_log )

    pos = 28 if int(ano) >= 2017 else 10
    versao = '' if int(ano) >= 2017 else '_Antigo'
    tipo = lst_arquivos[0][pos]
    dic_tipo = {}
    dic_tipo['M'] = 'MESTRE'
    dic_tipo['I'] = 'ITEM'
    #### ALT001 - Inicio - Adequar script para carregar CONTROLES
    dic_tipo['D'] = 'DESTINATARIO'
    dic_tipo['C'] = 'CONTROLE'
    #### ALT001 - Fim

    modelo_ctl = 'CNV115_%s%s.ctl'%( dic_tipo[tipo], versao )
    arq_ctl = '%s_%s_%s_%s.ctl'%( serie, ano, mes, tipo )
    if not os.path.isfile( os.path.join(dir_execucao,modelo_ctl) ) :
        log('Erro falta arquivo de modelo (.ctl) : ', modelo_ctl)
        return False 

    retorno = True
    for arq in lst_arquivos :
        ##### Monta o novo arquivo .ctl
        fd = open(os.path.join(dir_execucao,modelo_ctl), 'r')
        fd_ctl = open('./%s'%(arq_ctl), 'w')
        # vol = int(arq.split('.')[-1])
        vol = arq.split('.')[-1]
        for lin in fd :
            if lin.__contains__('$$UF$$') :
                lin = lin.replace('$$UF$$', uf )
            elif lin.__contains__('$$VOL$$') :
                # lin = lin.replace('$$VOL$$', str(vol) )
                lin = lin.replace('$$VOL$$', vol )
            elif lin.__contains__('$$ID$$') :
                lin = lin.replace('$$ID$$', str(id_serie) )
            #### ALT001 - Inicio - Adequar script para carregar CONTROLES
            elif lin.__contains__('$$SERIE$$') :
                lin = lin.replace('$$SERIE$$', str(serie) )
            #### ALT001 - Fim
            elif lin.__contains__('$$IN$$') :
                lin = lin.replace('$$IN$$', os.path.join( diretorio, arq ) )


            fd_ctl.write(lin)

        fd.close()
        fd_ctl.close()
        
        ## Executa o loader.
        # comando = """sqlldr openrisow/openrisow@%s %s data=\'%s\' > log/%s.log"""%( variaveis['banco'], os.path.join( os.path.realpath('.'), 'ctl', arq_ctl ), os.path.join(diretorio, arq.replace(' ', '\\ ')), arq_ctl[:-4] )
        # comando = """sqlldr gfcarga/vivo2019@%s %s > log/%s.log"""%( variaveis['banco'], os.path.join( os.path.realpath('.'), 'ctl', arq_ctl ), arq_ctl[:-4] )
        # comando = """sqlldr gfcarga/vivo2019@%s %s > log.log"""%( configuracoes.banco, os.path.join( os.path.realpath('.'), 'ctl', arq_ctl ) )
        
        comando = """sqlldr %s/%s@%s %s > log.log"""%(configuracoes.userBD, configuracoes.pwdBD, configuracoes.banco, arq_ctl )
        os.system(comando)

        ## Buscando os erros ocorridos.
        txt = "-"*100
        txt += """
Carregando arquivo : %s
"""%(arq)
        os.system("cat %s.log | grep -e 'ORA-' -e 'Rejected' -e 'SQL.Loader-' | head -20 > Resumo_%s.log"%(arq_ctl[:-4], arq_ctl[:-4])) 
        fd = open("Resumo_%s.log"%(arq_ctl[:-4]), 'r' )
        erros = fd.readlines()
        fd.close()
        if len(erros) > 0 :
            txt += """
Erros encontrados :
   """
            txt += '   '.join( x for x in erros )
            retorno = False

        ## Gerando resumo da execucao.
        os.system("cat %s.log | grep -e Row -e 'Total logical records' > Resumo_%s.log"%(arq_ctl[:-4], arq_ctl[:-4])) 
        fd = open("Resumo_%s.log"%(arq_ctl[:-4]), 'r' )
        resumo = fd.readlines()
        fd.close()
        txt += """
******************************************************
*** Resumo da execução : ( %s )
*** """%(arq)
        txt += '*** '.join( x for x in resumo )
        txt += '******************************************************\n'

        txt += "-"*100
        txt += "\n"
        log(txt)

        try :
            os.remove( "Resumo_%s.log"%(arq_ctl[:-4]) )
        except :
            pass
    time.sleep(1)
    if not retorno :
        os.system( '> ERRO_%s'%(pid) )
    return retorno   
        
def encodingDoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r', encoding='iso-8859-1')
        t = fd.read()
        fd.close()
    except :
        return 'utf-8'
    return 'iso-8859-1'

if __name__ == "__main__":
    name_script = os.path.basename(sys.argv[0]).replace('.py', '')
    log("-"*150)
    ret = 0
    txt = ''
    try:
        comum.carregaConfiguracoes(configuracoes)
        params = {}
        if len(sys.argv) > 4 :
            if '-DT' not in sys.argv or '-UF' not in sys.argv :
                ret = 97
            else :
                p = ''
                print(sys.argv)
                for i in range(1,len(sys.argv)) :
                    print(i,'->', sys.argv[i])
                    if sys.argv[i].startswith('-') :
                        p = sys.argv[i]
                    else :
                        if p :
                            if params.get(p, False) :
                                params[p] += ' '
                            params[p] = params.get(p,'') + sys.argv[i]
                        else :
                            ret = 98

        else :
            ret = 99

        variaveis['DT_REF']     = params.get('-DT', False)
        variaveis['UF']         = params.get('-UF', False)
        variaveis['Serie']      = params.get('-S', False)
        variaveis['Processar']  = params.get('-P', 'Todos').upper()
                
        log('UF a processar ..............: %s'%( variaveis['UF'] ))
        log('Data referencia a processar .: %s'%( variaveis['DT_REF'] ))
        log('Serie(s) a processar ........: %s'%( variaveis['Serie'] ))

        if not ret :
            try :
                mes, ano = variaveis['DT_REF'].split('/')
                datetime.datetime( int(ano), int(mes), 1 )
            except :
                log('Erro no parametro de -DT conjunto mes/ano passado não é um mes válido.')
                ret = 96
            
        if ret :
            log('ERRO - Erro nos parametros passados para o script.')
            log('')
            log('Exemplo de execução :')
            log('     ./%s.py -UF SP -DT 01/2015 -S UK, ASS, C'%(name_script))
            log('')
            log('Parametros :')
            log('   -UF < Sigla da uf > ')
            log('   -DT < Periodo de dados ex: 03/2016 >')
            log('   -S  <lista de series a processar ( Opcional )>')
            log('           A serie é opcional, caso não informado considera todas. ')
            log('------------------------------------------------------------------------')
            log('')
        else :
            if not ret :
                try:
                    ret = processar()
                except Exception as e:
                    txt = traceback.format_exc()
                    log(str(e))
                    raise Exception(str(e))
                        
    except Exception as e:
        txt = traceback.format_exc()
        log("ERRO = " + str(e))
        ret = 2
    
    log("="*150)
    sts = 'SUCESSO' if ret == 0 else 'ERRO'
    log( "Status Final ...: " + sts )
    sys.exit(ret)
