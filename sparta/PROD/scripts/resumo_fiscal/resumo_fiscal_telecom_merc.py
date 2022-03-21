#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..:
  MODULO ...:
  SCRIPT ...: resumo_fiscal
  CRIACAO ..: 15/12/2020
  AUTOR ....: Airton Borges da Silva Filho / KYROS Consultoria
  DESCRICAO :
----------------------------------------------------------------------------------------------
  HISTORICO :
              20210903 - Padronização para o novo Painel de Execuções.
----------------------------------------------------------------------------------------------
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
comum.log.gerar_log_em_arquivo = False
comum.carregaConfiguracoes(configuracoes)
#### PATRONIZACAO PARA O PAINEL DE EXECUCOES....

import datetime
import re
import subprocess

def ultimodia(ano,mes):
   return(31 if mes == 12 else datetime.date.fromordinal((datetime.date(ano,mes+1,1)).toordinal()-1).day)

def parametros():
    global ret
    ret = 0

#### Recebe, verifica e formata os argumentos de entrada.

    if (len(sys.argv) == 4
        and len(sys.argv[1])==6
        and int(sys.argv[1][:2])>0
        and int(sys.argv[1][:2])<13
        and int(sys.argv[1][2:])<=datetime.datetime.now().year
        and int(sys.argv[1][2:])>(datetime.datetime.now().year)-25
        ):

        mesanoi = sys.argv[1].upper()
        mesi  = sys.argv[1][:2].upper()
        anoi  = sys.argv[1][2:].upper()
        datai = "01"+mesi+anoi
        dataf = str(ultimodia(int(anoi),int(mesi)))+str(mesi)+str(anoi)
        sn=sys.argv[3].upper()

        iei=sys.argv[2].upper()
        iei = re.sub('[^0-9]','',iei)
        if ( (iei == "") or (iei == "''") or (iei == '""') or (int("0"+iei) == 0)):
            iei = "*"

    else :
        log("-" * 100)
        log("#### ")
        log('#### ERRO - Erro nos parametros do script.')
        log("#### ")
        log('#### Exemplo de como deve ser :')
        log('####      %s <MMAAAA> <IE> '%(sys.argv[0]))
        log("#### ")
        log('#### Onde')
        log('####      <mes e ano> = MMAAAA ex: 112020')
        log('####      <IE>        = Inscição Estadual no formato 999999999999 ex: 108383949112')
        log('####      <S/N>       = Informe N para somente Telecom.  Informe S para Telecom e Mercadoria.')
        log("#### ")
        log('#### ')
        log("-" * 100)
        log("")
#        log("Retorno = 99")
        ret = 99
        return(False,False,False,False)

    return(datai, dataf, iei, sn)

if __name__ == "__main__":
    global ret
    log('#'*100)
    log("# ")
    log("# - INICIO - RESUMO FISCAL TM")
    log("# ")
    log('#'*100)
    ret = 0
    retorno = parametros()
    if (retorno[0] != False):
        ret     = 0
        dataii  = retorno[0]
        datafi  = retorno[1]
        iei     = retorno[2]
        sn      = retorno[3]
           
        if (sn != "SOMENTE TELECOM" or sn == "S"):
            tipores = "Resumo de TELECOM E MERCADORIAS"
            comando = '/arquivos/JAVACorretto_11/bin/java -Xms512m -jar /arquivos/java/mastersaf-gf-1.0.jar RESUMO_FISCAL OPEN TBRA "" '
            comando = comando + iei + " " + dataii + " " + datafi + " SM SM ST ST F"
        else:
            tipores = "Resumo somente de TELECOM"
            comando = '/arquivos/JAVACorretto_11/bin/java -Xms512m -jar /arquivos/java/mastersaf-gf-1.0.jar RESUMO_FISCAL OPEN TBRA "" '
            comando = comando + iei + " " + dataii + " " + datafi + " ST ST F"
            
        
        log("Opções a serem executadas:")
        log("Tipo de resumo ... = ", sn)
        log("Inscrição Estadual = ", iei)
        log("Data inicial ..... = ", dataii)
        log("Data final ....... = ", datafi)

      
        log("-"*100)
        log("# Executando o comando :")
        log("# ",comando)
        log("-"*100)

        diratual = os.getcwd()
        direxec = "/arquivos/AUDITORIA"
        os.chdir(direxec)
        ret = subprocess.call(comando, shell=True)
        os.chdir(diratual)
    else:
        ret = 99

    if (ret != 0):
        log("-" * 100)
        log('#### ERRO - RETORNO = ', ret)
        log("-" * 100)
        
    log('#'*100)
    log("# ")
    log("# - FIM - RESUMO FISCAL TM")
    log("# ")
    log("#"*100)        

    if (ret > 0):
        log("ERRO - Verifique as mensagens anteriores...")

    log("Codigo de saida = ",ret)
    sys.exit(ret)




