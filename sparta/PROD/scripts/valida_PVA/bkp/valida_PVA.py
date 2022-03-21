#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: GF
  MODULO ...:
  SCRIPT ...: valida_PVA.py
  CRIACAO ..: 09/12/2021
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
    * 09/12/2021 - Welber Pena de Sousa - Kyros Tecnologia
        - Criacao do script.
----------------------------------------------------------------------------------------------
    
----------------------------------------------------------------------------------------------
"""

import sys
import os
import shutil
import calendar
import glob
import cx_Oracle
import time

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

    if ret == 0 : 
        if not IniciaExecucaoPVArobo(configuracoes.id_serie) :
            ret = 9 

    log('-'*150) 
    
    return True if not ret else False


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
                           "       ,FP.fk_id_clone = 44                                        " + chr(10) +
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


if __name__ == "__main__":
    log("-"*100)
    log("  INICIO DO VALIDA PVA  ".center(120,'#'))
    comum.carregaConfiguracoes(configuracoes)
    comum.addParametro('ID_SERIE', None, "ID da serie e ser processada.", True)
    ret = 0 

    if not comum.validarParametros():
        ret = 1
    else:
        configuracoes.id_serie = comum.getParametro('ID_SERIE')
        try :
            if not processar():
                ret = 2
        except Exception as e :
            log('ERRO ao processar:', e)
            raise e
        # if (ret > 0):
        log("### Retorno da execução ..:", ret)

    log("  FIM DO VALIDA PVA  ".center(120,'#'))
    
    sys.exit(ret)


