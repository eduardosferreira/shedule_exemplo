#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: GF
  MODULO ...:
  SCRIPT ...: geraConv115.py
  CRIACAO ..: 04/02/2020
  AUTOR ....: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO:
    - Busca na tabela GFCARGA.TSH_SERIE_LEVANTAMENTO os dados referentes ao ID_SERIE passado
        como parametro para o script.

    - Acessar diretório /ARQUIVOS/CONVENIO_115

    - Formatar linha de comando para geração do Convênio 115, conforme regra detalhada na
        tabela “REGRA LINHA COMNADO 115”:

    java -Xms512m -jar /arquivos/java/mastersaf-gf-1.0.jar CNV115 [Empresa] [filial] [inscricaoEstadual]
        [dataInicial] [dataFinal] [modeloDocumento] [serie] [usuario] [diretorio] [retificacao]
        [renumeraItens] [cofaturamento] [convenio52] [geraHash] [tipoPessoaNF] [quantidadeVolume]
        [versao] [geraTerminalFaturado] [geraFaturaServico] [geraCarregamentoCredito] [geracaoAvulsa]
        [diretorio201] [simplesConferencia]

    Exemplo:
        java -Xms512m -jar /arquivos/java/mastersaf-gf-1.0.jar CNV115 TBRA 0001 "" 01012015 31012015
        22 "TE" A0863658 /arquivos/CONVENIO_115/ T T F F T T 1000000 07/12 F F F F /arquivos/CONVENIO_115/ T

    - Executar linha de comando para geração do Convênio 115.

        - Caso execução concluída com sucesso:
            - Buscar na tabela TSH_QUEBRA_VOLUME, utilizando Empresa, Filial, Mes Ano e Série.
                - Caso registro encontrado, executar python redistribuiVolumes.py e atualizaControle115.py
                    sequencialmente para o ID da série .
                    (ambos estão contidos no diretório /arquivos/TESHUVA/scripts_rpa/RecalculoVolumes/)

                - Caso registro não encontrado, continuar com o processamento.

    - Chamar python posicionaPVA.py, contido no diretório /arquivos/TESHUVA/scripts_rpa/unificado,
        passando o id da série como parâmetro

    - Caso erro de execução, parar o processamento.

----------------------------------------------------------------------------------------------
  HISTORICO:
    * 04/02/2020 - Welber Pena de Sousa - Kyros Tecnologia
        - Criacao do script.

----------------------------------------------------------------------------------------------
    * 12/02/2020 - Welber Pena de Sousa - Kyros Tecnologia
        - Alterar a regra para definição se chama ou não o script de quebra, retirar a pesquisa
          na tabela TSH_QUEBRA_VOLUME e substituir pela logica abaixo:

            Execução da query abaixo:

                SELECT DISTINCT EMPS_COD, FILI_COD, SERIE, MES_ANO, COUNT QUANTIDADE
                FROM TSH_CONTROLE_ARQ_CONV_115[TAB]
                WHERE AREA = 'PROTOCOLADO'
                      AND SUBSTR(NOME_ARQUIVO,DECODE([POS]),1) = 'M'
                      AND MENSAGEM_STATUS = 'PROCESSADO'
                      AND SERIE = REPLACE([SERIE],’ ‘,’’)
                GROUP BY EMPS_COD, FILI_COD, SERIE, MES_ANO

            Onde:
                [SERIE] = série em processamento
                [POS] = se ano em processamento <= 2016 então 11 senão 29
                [TAB] = se ano em processamento = 2017 então 17, se 2016 então 16, se outro, então espaço em branco

        - Se o total de volumes gerado (considerar penas quantidade de registros tipo M) for
          diferente do campo quantidade retornado pela query, então executar quebra, senão,
          não executar quebra.

        - Se a query retornar notfound, forçar erro com mensagem:
            “Não foi carregado o arquivo protocolado para esta série. Favor realizar a carga e reprocessar”

----------------------------------------------------------------------------------------------
    * 18/02/2020 - Flavio Teixeira - ALT001
         - Tratamento de filial para anos anteriores a 2017.

----------------------------------------------------------------------------------------------
    * 04/03/2020 - Flavio Teixeira - ALT002
         - Ajuste parametro simples conferencia.

----------------------------------------------------------------------------------------------
    * 10/07/2020 - Welber Pena - ALT003 - Refeito 15/09/2020 - Fausto
         A equipe do Tributário informou que o Convênio 115 de Pernambuco -
        filial 3506 - deve ser gerado com Status N - Normal de Julho/15 a
        Dezembro/17.
        A partir de Janeiro/18 passam a ser gerados como S - Substituto.

----------------------------------------------------------------------------------------------
    * 17/03/2021 - Welber Pena
        Documentação : ALT005_Geracao_Conv115_parametrizavel.pdf
        ALT005 
            - Inseridos os campos INDICADOR_RETIFICACAO e SEQUENCIA na basca de dados da serie.
            - De acordo com o INDICADOR_RETIFICACAO, altera o parâmetro de retificação na linha de comando da geração do convênio (java).
            - Acertando a sequencia de retificação dos arquivos de acordo com o campo SEQUENCIA.

----------------------------------------------------------------------------------------------
    * 17/08/2021 - Welber Pena 
        ALT006
            - Script re-escrito para atender a nova estrutura do Painel de Execucoes .

----------------------------------------------------------------------------------------------
    * 26/10/2021 - Welber Pena 
        ALT007 - Parametro EXEC_PVA
            - Alteração para script receber o parametro de execucao do PVA.

----------------------------------------------------------------------------------------------
"""

import sys
import os
import shutil
import calendar
import glob
import cx_Oracle
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)

import configuracoes
import comum
import sql

log.gerar_log_em_arquivo = True

# configuracoes.dir_trabalho = '/arquivos/CONVENIO_115'
# configuracoes.diretorio_base_destino = '/portaloptrib/LEVCV115'


def processar():
    id_serie = configuracoes.id_serie
    ret = 0

    if ret == 0:
        configuracoes.dic_dados_serie = comum.buscaDadosSerie(id_serie)
        if not configuracoes.dic_dados_serie:
            ret = 4
    
    if ret == 0 and not executaJAVA():
        ret = 5
    
    if ret == 0 and not quebraVolume():
        ret = 7
    
    ### ALT005 - Inicio
    if ret == 0 and not acertaSequencia():
        ret = 8
    ## ALT005 - Fim

    ### ALT007 - Inicio
    if ret == 0 and configuracoes.exec_PVA.startswith('S') : 
        if not IniciaExecucaoPVArobo(configuracoes.id_serie) :
            ret = 9 
    ### ALT007 - Fim

    log('-'*150) 
    
    return True if not ret else False


def executaJAVA(): 
    """
    - Acessar diretório /ARQUIVOS/CONVENIO_115

    - Formatar linha de comando para geração do Convênio 115, conforme regra detalhada na
        tabela “REGRA LINHA COMNADO 115”:

    java -Xms512m -jar /arquivos/java/mastersaf-gf-1.0.jar CNV115 [Empresa] [filial] [inscricaoEstadual]
        [dataInicial] [dataFinal] [modeloDocumento] [serie] [usuario] [diretorio] [retificacao]
        [renumeraItens] [cofaturamento] [convenio52] [geraHash] [tipoPessoaNF] [quantidadeVolume]
        [versao] [geraTerminalFaturado] [geraFaturaServico] [geraCarregamentoCredito] [geracaoAvulsa]
        [diretorio201] [simplesConferencia]

    Exemplo:
        java -Xms512m -jar /arquivos/java/mastersaf-gf-1.0.jar CNV115 TBRA 0001 "" 01012015 31012015
        22 "TE" A0863658 /arquivos/CONVENIO_115/ T T F F T T 1000000 07/12 F F F F /arquivos/CONVENIO_115/ T

    - Executar linha de comando para geração do Convênio 115.

        - Caso execução concluída com sucesso:
            - Buscar na tabela TSH_QUEBRA_VOLUME, utilizando Empresa, Filial, Mes Ano e Série.
                - Caso registro encontrado, executar python redistribuiVolumes.py e atualizaControle115.py
                    sequencialmente para o ID da série .
                    (ambos estão contidos no diretório /arquivos/TESHUVA/scripts_rpa/RecalculoVolumes/)

                - Caso registro não encontrado, continuar com o processamento.
    """
    dir_trabalho = configuracoes.dir_trabalho
    log('Acessando diretorio', dir_trabalho)
    if not os.path.isdir(dir_trabalho):
        os.makedirs(dir_trabalho)
    os.chdir('/arquivos/CONVENIO_115')
    ano = configuracoes.dic_dados_serie['ano']
    mes = configuracoes.dic_dados_serie['mes']
    ultDiaDoMes = calendar.monthrange(int(ano), int(mes))[1]
    data_inicio_mes = '01%s%s' % (mes, ano)
    data_fim_mes = '%s%s%s' % (ultDiaDoMes, mes, ano)
    versao = '07/12' if int(ano) < 2017 else '60/15'

    if int(ano) >= 2020:
        versao = '29/18'

    ######  10/07/2020 - Welber Pena - ALT003
    ###  A equipe do Tributário informou que o Convênio 115 de Pernambuco -
    ###  filial 3506 - deve ser gerado com Status N - Normal de Julho/15 a
    ###  Dezembro/17.
    ###  A partir de Janeiro/18 passam a ser gerados como S - Substituto.

    retificacao = 'T'
    if str(configuracoes.dic_dados_serie['filial']) == '3506':
        if int('%s%s' % (ano, mes)) >= 201507:
            if int('%s%s' % (ano, mes)) < 201801:
                retificacao = 'F'
    ##### FIM ALT003

    ####### 17/03/2021 - Welber Pena - ALT005
    else:
        if configuracoes.dic_dados_serie['indicador_retificacao'].upper() != 'S':
            retificacao = 'F'
    ####### FIM ALT005

    # cmd_java = """/arquivos/JAVACorretto_11/bin/java -Xms512m -Xmx2g -jar -Dcom.sun.management.jmxremote.authenticate=false -Dcom.sun.management.jmxremote.ssl=false -Dcom.sun.management.jmxremote.port=3011 /arquivos/java/mastersaf-gf-1.0.jar CNV115 %s %s "" %s %s %s "%s" RBTESHGF10 %s/ %s T F F T T 1000000 %s F F F F %s/ "" F "" ""
# """ % ( configuracoes.dic_dados_serie['empresa'], configuracoes.dic_dados_serie['filial'], data_inicio_mes, data_fim_mes, configuracoes.dic_dados_serie['modelo_nf'], configuracoes.dic_dados_serie['serie_original'], dir_trabalho, retificacao, versao, dir_trabalho )
    cmd_java = """/arquivos/JAVACorretto_11/bin/java -Xms512m -jar /arquivos/java/mastersaf-gf-1.0.jar CNV115 %s %s "" %s %s %s "%s" RBTESHGF10 %s/ %s T F F T F 1000000 %s F F F F %s/ "" F "" ""
""" % ( configuracoes.dic_dados_serie['empresa'], configuracoes.dic_dados_serie['filial'], data_inicio_mes, data_fim_mes, configuracoes.dic_dados_serie['modelo_nf'], configuracoes.dic_dados_serie['serie_original'], dir_trabalho, retificacao, versao, dir_trabalho )
    
    
    ### Deve ser revisto com o Teixeira se sera feito com as duas linhas abaixo
    # config = obterConfiguracao()    
    movimentaObrigacao()
    ##################

    
    log('Comando JAVA : %s' % (cmd_java))
    log('  INICIO execucao do JAVA mastersaf-gf-1.0.jar  '.center(150,'='))
    r = os.system(cmd_java)
    log('  FIM da execucao do JAVA mastersaf-gf-1.0.jar  '.center(150,'='))
    # r = os.system('ls no*')
    # r = 0

    log('Resultado ...:', r)
    if r == 0:
        log(' -> Java executado com SUCESSO !')
        os.system("chmod g+x %s/*" % (dir_trabalho))
    else:
        log(' -> Java executado com ERRO !')
        raise Exception('Conv115', 'Erro no Java de geracao de conv115')
        return False

    #########################################################################################################################
    ###### Alteracao : 12/02/2020 - Welber Pena de Sousa - Kyros Tecnologia
    # - Alterar a regra para definição se chama ou não o script de quebra, retirar a pesquisa
    #   na tabela TSH_QUEBRA_VOLUME e substituir pela logica abaixo:

    #     Execução da query abaixo:

    #         SELECT DISTINCT EMPS_COD, FILI_COD, SERIE, MES_ANO, COUNT QUANTIDADE
    #         FROM TSH_CONTROLE_ARQ_CONV_115[TAB]
    #         WHERE AREA = 'PROTOCOLADO'
    #               AND SUBSTR(NOME_ARQUIVO,DECODE([POS]),1) = 'M'
    #               AND MENSAGEM_STATUS = 'PROCESSADO'
    #               AND SERIE = REPLACE([SERIE],’ ‘,’’)
    #         GROUP BY EMPS_COD, FILI_COD, SERIE, MES_ANO

    #     Onde:
    #         [SERIE] = série em processamento
    #         [POS] = se ano em processamento <=2016 então 11 senão 29
    #         [TAB] = se ano em processamento = 2017 então 17, se 2016 então 16, se outro, então espaço em branco

    # - Se o total de volumes gerado (considerar apenas quantidade de registros tipo M) for
    #   diferente do campo quantidade retornado pela query, então executar quebra, senão,
    #   não executar quebra.

    # - Se a query retornar notfound, forçar erro com mensagem:
    #     “Não foi carregado o arquivo protocolado para esta série. Favor realizar a carga e reprocessar”

    #ALT001 - Inicio
    #ALT006 - Inicio
    #executaPVA() 
    
    ### ALT007 - Inicio
    if configuracoes.exec_PVA.upper().startswith('S') :
        if not atualizaStatusRobo( configuracoes.id_serie, 22) :
            return False
    ### ALT007 - Fim
    
    movimentaObrigacao()

    #ALT006 - Fim
    #ALT001 - Fim
    qt_arqs_gerados = 0
    log('Contando arquivos gerados ...')
    pos = 10 if int(ano) < 2017 else 28

    #ALT001 - Inicio
    if os.path.isdir(os.path.join(configuracoes.dic_dados_serie['dir_serie'], 'OBRIGACAO')) :
        for arq in os.listdir(os.path.join(configuracoes.dic_dados_serie['dir_serie'], 'OBRIGACAO')):
            if len(arq) > pos and arq[pos] == 'M':
                #ALT003 - Inicio
                #if  int(ano) >= 2017:
                #    v_modelo = arq[16:18]
                #    v_serie = arq[18:21]
                #    if  v_modelo == '21' or v_serie == 'TE ':
                #        chamaPaliativoTerminal(arq, configuracoes.dic_dados_serie['dir_serie'], 'LayoutMestre.csv')
                #ALT003 - Fim
                qt_arqs_gerados += 1

    #ALT001 - for arq in os.listdir(dir_trabalho):
    #ALT001 -     log('arq = ' + arq)
    #ALT001 -     if arq.startswith('SP'):
    #ALT001 -          if len(arq) > pos and arq[pos] == 'M':
    #ALT001 -             if arq.__contains__(configuracoes.dic_dados_serie['serie']):
    #ALT001 -                if arq[-12:].startswith('%s%s' % (ano[-2:],mes)):
    #ALT001 -                     qt_arqs_gerados += 1

    log('Foram gerados .....: %s arquivo(s)' % ( qt_arqs_gerados ))

    log('Buscando na base de dados a quantidade de arquivos protocolados.')

    ## ALT004 - cmd_sql = """
    ## ALT004 - select distinct emps_cod, fili_cod, serie, MES_ANO, count(*) quantidade
    ## ALT004 - from TSH_CONTROLE_ARQ_CONV_115%s
    ## ALT004 - where area = 'PROTOCOLADO'
    ## ALT004 -     and substr(nome_arquivo,%s,1) = 'M'
    ## ALT004 -     and MENSAGEM_STATUS = 'Processado'
    ## ALT004 -     and serie = replace('%s',' ','')
    ## ALT004 -     and MES_ANO = to_date('01/%s/%s', 'dd/mm/yyyy' )
    ## ALT004 -     and fili_cod = '%s'
    ## ALT004 - group by emps_cod, fili_cod, serie, MES_ANO
    ## ALT004 - """ % ("" if int(ano) not in (2016,2017) else ano[-2:], pos+1, configuracoes.dic_dados_serie['serie'], mes, ano, configuracoes.dic_dados_serie['filial'] if int(ano) >= 2017 else '0' )

    #ALT001 """ % ("" if int(ano) not in (2016,2017) else ano[-2:], pos+1, configuracoes.dic_dados_serie['serie'], mes, ano, configuracoes.dic_dados_serie['filial'] )

    ### ALT004 - Inicio
    ## Verificar a quantidade de volumes protocolado na tabela TSH_CONTROLE_ARQ_CONV_115 na base GFREAD.
    ## - Deve-se considerar os parametros ctr_apur_dtini, ctr_ser_ori, fili_cod, trazendo o max(ctr_volume).
    cmd_sql = """
    SELECT max(to_number(ctr_volume))
    FROM  openrisow.CTR_IDENT_CNV115
    where ctr_apur_dtini = to_date('%s', 'dd/mm/yyyy')
            and ctr_ser_ori = '%s'
            and fili_cod = '%s'
    """ % (configuracoes.dic_dados_serie['data_ini_apuracao'], configuracoes.dic_dados_serie['serie_original'], configuracoes.dic_dados_serie['filial'] )
    ### ALT004 - FIM

    obj_sql = sql.geraCnxBD(configuracoes) 
    
    obj_sql.executa( cmd_sql )
    linha = obj_sql.fetchone()

    log(linha)

    if linha[0] is None:

        msgErro = (
            "Tabela de Controle (CTR_IDENT_CNV115) e Diretorio Protocolado nao possui registro para esta serie: "
            + " Data Apuracao:" + configuracoes.dic_dados_serie['data_ini_apuracao']
            + " Serie: " + configuracoes.dic_dados_serie['serie_original']
            + " Filial: " + configuracoes.dic_dados_serie['filial']
        )
        qtd_ret = 0
        dir_protocolados_verificar = os.path.join( configuracoes.dic_dados_serie['dir_serie'].replace('/OBRIGACAO', ''), 'PROTOCOLADO' )
        if os.path.isdir(dir_protocolados_verificar):
            for item_arq in os.listdir(dir_protocolados_verificar):
                if item_arq[pos] == 'C':
                    qtd_ret += 1
            
        if qtd_ret == 0:
            raise Exception(msgErro)
    else:
        ## ALT004 - qtd_ret = linha[4]
        qtd_ret = linha[0]

    configuracoes.dic_dados_serie['quebra_volumes'] = True if qt_arqs_gerados != int(qtd_ret) else False

    #########################################################################################################################

    return True

### ALT007 - Inicio

def atualizaStatusRobo(id_serie, status):
    log("Atualizando status do ROBO ...")

    conexao = obterConexaoRobo()
    
    idFilaPeriodo = 0
    
    try:
        cursor = conexao.cursor()
        try:
            log(" ")    
            log("-------------------------------------------------------------------------")    
            log("- Iniciando Atualizacao do Banco RPA para posicionar na etapa de PVA...")    
            log(" ")    
            log("- id_serie         = ", id_serie)        
            log(" ")    
            log("- Atualizando Filas  invalidas (Sem data Fim)...")    
            log(" ")    
            
            cursor.execute(" UPDATE FILA_PERIODO FP                                            " + chr(10) +
                           "    SET FP.FIM  = current_timestamp                                " + chr(10) +
                           "       ,FP.ERRO = DECODE(FP.ETAPA,'60','Y'                         " + chr(10) +
                           "                                 , '3','Y'                         " + chr(10) +
                           "                                 ,'61','Y'                         " + chr(10) +
                           "                                 , '6','Y'                         " + chr(10) +
                           "                                 ,'47','Y'                         " + chr(10) +
                           "                                 ,'48','Y'                         " + chr(10) +
                           "                                 ,'53','Y'                         " + chr(10) +
                           "                                 ,'57','Y'                         " + chr(10) +
                           "                                 ,'N'                              " + chr(10) +
                           "                                 )                                 " + chr(10) +
                           " WHERE FP.ROWID IN (                                               " + chr(10) +
                           "                    SELECT FPS.ROWID                               " + chr(10) +
                           "                      FROM FILA_PERIODO FPS                        " + chr(10) +
                           "                          ,FILA         FL                         " + chr(10) +
                           "                     WHERE FL.ID                    = FPS.ID_FILA  " + chr(10) +
                           "                       AND FPS.FIM                  IS NULL        " + chr(10) +
                           "                       AND FL.ID_SERIE_LEVANTAMENTO = :ID_SERIE    " + chr(10) +
                           "                   )                                               "
                          ,(id_serie,))
                                          
            conexao.commit()

            log("- Inserindo registro de reposicionamento no status ", status)    
            log(" ")    
                
            cursor.execute(" INSERT                                   " + chr(10) +
                           "   INTO FILA_PERIODO(ID_FILA              " + chr(10) +
                           "                    ,USUARIO_ID           " + chr(10) +
                           "                    ,STATUS               " + chr(10) +
                           "                    ,ETAPA                " + chr(10) +
                           "                    ,INICIO               " + chr(10) +
                           "                    ,FIM                  " + chr(10) +
                           "                    ,ERRO                 " + chr(10) +
                           "                    )                     " + chr(10) +
                           " SELECT DISTINCT ID                       " + chr(10) +
                           "                ,1                        " + chr(10) +
                           "                ,:STATUS                  " + chr(10) +
                           "                ,:STATUS                  " + chr(10) +
                           "                ,current_timestamp        " + chr(10) +
                           "                ,current_timestamp        " + chr(10) +
                           "                ,'N'                      " + chr(10) +
                           "   FROM FILA                              " + chr(10) +
                           "  WHERE ID_SERIE_LEVANTAMENTO = :ID_SERIE "
                          ,(status,status,id_serie))
                     
            conexao.commit()
        
        finally:
            cursor.close()
            
    finally:
        conexao.close()
   
    log("- Tratativas concluidas com sucesso! ")    
    log("-------------------------------------------------------------------------")    
    log(" ")

    return True


def obterConexaoRobo():
    conexao = cx_Oracle.connect("portaloptrib/portaloptrib_4234@BANCORPA", threaded=True)
    conexao.autocommit = False
    return conexao


def IniciaExecucaoPVArobo(id_serie) :
    log('Preparando a execucao do IniciaExecucaoPVArobo ... Aguarde ...')
    log(' - Iniciando o IniciaExecucaoPVArobo ...')

    conexao = obterConexaoRobo()
    try:
        cursor = conexao.cursor()
        try:
            log(" ")    
            log("-------------------------------------------------------------------------")    
            log("- I N I C I A -  P  V  A --- --- ---")    
            log(" ")    
            log("- id_serie         = ", id_serie)        
            log(" ")               
            log("- Inserindo registro de start de PVA")    
            log(" ")    
                
            cursor.execute(" INSERT                                   " + chr(10) +
                           "   INTO FILA_PERIODO(ID_FILA              " + chr(10) +
                           "                    ,USUARIO_ID           " + chr(10) +
                           "                    ,STATUS               " + chr(10) +
                           "                    ,ETAPA                " + chr(10) +
                           "                    ,INICIO               " + chr(10) +
                           "                    )                     " + chr(10) +
                           " SELECT DISTINCT ID                       " + chr(10) +
                           "                ,1                        " + chr(10) +
                           "                ,45                       " + chr(10) +
                           "                ,45                       " + chr(10) +
                           "                ,current_timestamp        " + chr(10) +
                           "   FROM FILA                              " + chr(10) +
                           "  WHERE ID_SERIE_LEVANTAMENTO = :ID_SERIE "
                          ,(id_serie,))
                     
            conexao.commit()
        except :
            log('Erro ao executar o script < IniciaExecucaoPVArobo.py > ... Abortando ...')
            return False
        finally:
            cursor.close()
            
    finally:
        conexao.close()
   
    log("- Tratativas concluidas com sucesso! ")    
    log("-------------------------------------------------------------------------")    
    
    return True

### ALT007 - Fim


def movimentaObrigacao():
    id_serie = configuracoes.id_serie
    log("Movimentando obrigacao ...")

    obj_sql = sql.geraCnxBD(configuracoes)
    obj_sql.executa(
            "SELECT UNFE_SIG, TO_CHAR(l.MES_ANO, 'YYMM'), TO_CHAR(l.MES_ANO, 'YY') ANO, REPLACE(SERIE,' ',''), f.UNFE_SIG||'/'||TO_CHAR(l.MES_ANO, 'YY/MM')||'/'||l.EMPS_COD||'/'||l.FILI_COD||'/SERIE/'||l.ID_SERIE_LEVANTAMENTO"+chr(10)+
            "FROM gfcarga.TSH_SERIE_LEVANTAMENTO l"+chr(10)+
            "  INNER JOIN OPENRISOW.FILIAL f ON l.EMPS_COD = f.EMPS_COD AND l.FILI_COD = f.FILI_COD"+chr(10)+
            "WHERE l.ID_SERIE_LEVANTAMENTO = :ID_SERIE",
            (id_serie,)
        )
    
    linha = obj_sql.fetchone()
    uf               = linha[0]
    ano_mes          = linha[1]
    ano              = int(linha[2])
    serie            = linha[3]
    diretorioDestino =  os.path.join( configuracoes.diretorio_base_destino, linha[4], "OBRIGACAO" )
    diretorioPVA     =  os.path.join( configuracoes.diretorio_base_destino, linha[4], "PVA" )

    log("-------------------------------------------------------------------------")    
    log("- Iniciando Movimentacao do Convenio 115...")    
    log(" ")    
    log("- configuracoes.dir_trabalho  = ", configuracoes.dir_trabalho)    
    log("- diretorioDestino = ", diretorioDestino)    
    log("- serie            = ", serie)    
    log("- ano              = ", str(ano))    
    log("- uf               = ", uf )    
    log("- ano_mes          = ", ano_mes)    
    log("-------------------------------------------------------------------------")    
    
    if  not os.path.exists(diretorioPVA):
        log("Criando Diretorio PVA...")
        os.makedirs(diretorioPVA)
        log("- Diretorio Criado!")
        log(" ")
        
    listaArquivos = [f for f in glob.glob(configuracoes.dir_trabalho + "/*", recursive=False)]    

    if os.path.isdir( diretorioDestino ):
        log("Limpando diretorio de destino de arquivos .:", diretorioDestino)
        for a in os.listdir(diretorioDestino):
            log(" - Apagando o arquivo .:", a)
            os.remove(os.path.join(diretorioDestino, a))

    if  ano <= 16:
        for arq in listaArquivos:
            nome_arquivo = arq.split(SD)[-1]
            
            if  os.path.isfile(arq):
                Arquivo_serie        = nome_arquivo[2:5].strip()
                Arquivo_ano_mes      = nome_arquivo[5:9]
                Arquivo_uf           = nome_arquivo[0:2]
                               
                if  Arquivo_serie   == serie and Arquivo_ano_mes == ano_mes and Arquivo_uf == uf:
                    log(" ")
                    log("-------------------------------------------------------------------------")
                    log("- Movendo arquivo... ")
                    log("- DE: " + arq )
                    log("- PARA: " + diretorioDestino + "/" + nome_arquivo)
                    log("  ")
                    shutil.move(arq, os.path.join(diretorioDestino, nome_arquivo) )
                    log("- Garantindo Permissao...")
                    os.chmod(diretorioDestino + "/" + nome_arquivo, 0o777)
                    log("- Permissao Concedida!")
                    
                    log("  ")
                    log("- Arquivo Movido Com sucesso! ")
                    log("-------------------------------------------------------------------------")
    else:
        for arq in listaArquivos:
            nome_arquivo = arq.split(SD)[-1]
            
            if  os.path.isfile(arq):
                Arquivo_serie        = nome_arquivo[18:21].strip()
                Arquivo_ano_mes      = nome_arquivo[21:25]
                Arquivo_uf           = nome_arquivo[0:2]
                                
                if  Arquivo_serie   == serie and Arquivo_ano_mes == ano_mes and Arquivo_uf == uf:
                    log(" ")
                    log("-------------------------------------------------------------------------")
                    log("- Movendo arquivo... ")
                    log("- DE: " + arq )
                    log("- PARA: " + diretorioDestino + "/" + nome_arquivo)
                    log(" ")
                    shutil.move(arq, os.path.join(diretorioDestino, nome_arquivo)  )
                    log("- Garantindo Permissao...")
                    os.chmod(diretorioDestino + "/" + nome_arquivo, 0o777)
                    log("- Permissao Concedida!")
                                         
                    log(" ")
                    log("- Arquivo Movido Com sucesso! ")
                    log("-------------------------------------------------------------------------")


def quebraVolume():
    if configuracoes.dic_dados_serie['quebra_volumes']:
        # - Caso registro encontrado, executar python redistribuiVolumes.py e atualizaControle115.py
        #             sequencialmente para o ID da série .
        #             (ambos estão contidos no diretório /arquivos/TESHUVA/scripts_rpa/RecalculoVolumes/)
        log('Iniciando o script < redistribuiVolumes.py > .... Aguarde ...')
        os.chdir( os.path.join(dir_base, 'scripts', 'Convenio115', 'redistribuir_volumes_conv115') )
        cmd_redistribui = './redistribuiVolumes.py %s' % ( configuracoes.dic_dados_serie['id_serie'] )
        log('  INICIO execucao do script redistribuiVolumes.py  '.center(150,'='))
        res = os.system(cmd_redistribui)
        log('  FIM da execucao do script redistribuiVolumes.py  '.center(150,'='))
        if res > 1:
            log('Erro ao executar o script < redistribuiVolumes.py > ... Verifique ... Abortando !!!!')
            return False
        log(' -> Script < redistribuiVolumes.py > executado com SUCESSO !')

        #### Comentado pois deve ser executado somente apos o PVA
        # log('Iniciando o script < atualizaControle115.py > .... Aguarde ...')
        # cmd_atualiza = './atualizaControle115.py %s' % ( configuracoes.dic_dados_serie['id_serie'] )
        # log('  INICIO execucao do script atualizaControle115.py  '.center(150,'='))
        # res = os.system(cmd_atualiza)
        # log('  FIM da execucao do script atualizaControle115.py  '.center(150,'='))
        # if res > 1:
        #     log('Erro ao executar o script < atualizaControle115.py > ... Verifique ... Abortando !!!!')
        #     return False
        # log(' -> Script < atualizaControle115.py > executado com SUCESSO !')

    else:
        log('Esta serie nao deve ser executado o scritp de < Recalculo de volumes > ... Pulando etapa !!!')

    return True


####### 17/03/2021 - Welber Pena - ALT005
### Criada a funcao abaixo para atender o item 3 da demanda.
def acertaSequencia():
    log('Acertando a sequencia de retificação dos arquivos ...')
    # dir_regerados = os.path.join( configuracoes.dic_dados_serie['diretorioDestino'].replace('/OBRIGACAO', ''), 'REGERADO' )
    dir_regerados = os.path.join(configuracoes.dic_dados_serie['dir_serie'], 'OBRIGACAO' )
    ano = configuracoes.dic_dados_serie['ano']
    pos = 10 if int(ano) < 2017 else 28
    log('- Diretorio de arquivos ..: %s' % (dir_regerados))

    if int(configuracoes.dic_dados_serie['sequencia']) > 1:
        if int(ano) >= 2017:
            if os.path.isdir(dir_regerados):
                for item_arq in os.listdir(dir_regerados):
                    if len(item_arq) > pos and item_arq[pos] == 'C':
                        log('  - Deletando o arquivo ........: %s' % (item_arq))
                        os.remove( os.path.join(dir_regerados, item_arq) )
                    elif len(item_arq) > pos and item_arq[pos] in [ 'I', 'M', 'D' ]:
                        novo_nome = item_arq[:pos-2] + str(configuracoes.dic_dados_serie['sequencia']).rjust(2,'0') + item_arq[pos:]
                        log('  - Alterando nome do arquivo ..: %s  para >>  %s' % (item_arq, novo_nome))
                        shutil.move( os.path.join(dir_regerados, item_arq), os.path.join(dir_regerados, novo_nome) )

            else:
                log(' - Diretorio de arquivos REGERADOS não encontrado.')
                return False
        else:
            log(' - Para series anteriores a 2017 nao é alterada a sequencia de retificação .')
            return False
    else:
        log(' - A sequencia de retificação não será alterada nesta execução, conforme parametros passados.')

    return True
####### FIM - ALT005


if __name__ == "__main__":
    log("-"*100)
    log("  INICIO DA GERACAO DO CONVENIO 115  ".center(120,'#'))
    comum.carregaConfiguracoes(configuracoes)
    comum.addParametro('ID_SERIE', None, "ID da serie e ser processada.", True)
    comum.addParametro('EXEC_PVA', None, "Executar PVA ? [S/n].", False, 'N', 'N')
    ret = 0 

    if not comum.validarParametros():
        ret = 1
    else:
        configuracoes.id_serie = comum.getParametro('ID_SERIE')
        configuracoes.exec_PVA = comum.getParametro('EXEC_PVA').upper()
        try :
            if not processar():
                ret = 2
        except Exception as e :
            log('ERRO ao processar:', e)
            raise e
        # if (ret > 0):
        log("### Retorno da execução ..:", ret)

    log("  FIM DA GERACAO DO CONVENIO 115  ".center(120,'#'))
    
    sys.exit(ret)


