#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-

"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: registro_1400_chamador.py
  CRIACAO ..: 27/08/2021
  AUTOR ....: Airton Borges da Silva Filho / KYROS Consultoria
  DESCRICAO :

            Este script recebe todos os parametros necessarios a execucao dos scripts de registro_1400 para os estados(UFs) existentes.

            VARIÁVEIS:
                <UF>     =  Estado (Unidade Federativa) ex: RJ
                <MMAAAA> =  MM: mes, numérico com dois dígitos. AAAA: ano, numérico com quatro dígitos.  ex: 012021
                <IE>     =  Inscriçao estadual. numérico.  ex: 77452443
    10/03/2022  - Eduardo da Silva Ferreira - Kyros Tecnologia
                - [PTITES-1688] Padrão de diretórios do SPARTA                 
               
----------------------------------------------------------------------------------------------
"""
import sys
import os
global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)
import configuracoes
import comum
import sql

import subprocess
import datetime
import time
import cx_Oracle
import glob
import shutil
import re
from pathlib import Path
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl import Workbook

global variaveis
global db
global arquivo_destino



def validauf(uf):
    return(True if (uf.upper() in ('AC','AL','AM','AP','BA','CE','DF','ES','GO','MA','MG','MS','MT','PA','PB','PE','PI','PR','RJ','RN','RO','RR','RS','SC','SE','SP','TO')) else False)


def parametros():
    global ret
    ufi = ""
    mesanoi = ""
    iei = "*"
    mesi = ""
    anoi = "" 
    ret = 0
    ufi = ""
  
    if (len(sys.argv) == 5
        
        and validauf(sys.argv[1])
        and len(sys.argv[2])==6  
        and int(sys.argv[2][:2])>0 
        and int(sys.argv[2][:2])<13
        and int(sys.argv[2][2:])<=datetime.datetime.now().year
        and int(sys.argv[2][2:])>2010
        ):
        
        iei=sys.argv[3].upper()
        iei = re.sub('[^0-9]','',iei)
        if ( (iei == "") or (iei == "''") or (iei == '""') or (int("0"+iei) == 0)):
            iei = "*"
            
        invai=sys.argv[4].upper()
             
    else :
        log("-" * 100)
        log("#### ")
        log('#### ERRO - Erro nos parametros do script.')
        log("#### ")
        log('#### Exemplo de como deve ser :')
        log('####      %s <UF> <MMAAAA> <IE>'%(sys.argv[0]))
        log("#### ")
        log('#### Onde')
        log('####      <UF>     = Estado. Ex: RJ')
        log('####      <MMAAAA> = mês e ano. Ex: Para junho de 2020 informe 062020')
        log('####      <IE>     = Inscição Estadual. Ex: 77452443')
        log('####      <S/N>    = Atualizar/Inserir registros na tabela INVA ?  Sim ou não (S/N) Válido só para RJ')
        log('#### OU')
        log('####      <UF>     = Estado. Ex: SP')
        log('####      <MMAAAA> = mês e ano. Ex: Para junho de 2020 informe 062020')
        log('####      <IE>     = Inscição Estadual. Ex: 108383949112')
        log('####      <S/N>    = Atualizar/Inserir registros na tabela INVA ?  Sim ou não (S/N) Válido só para RJ. Para SP qualquer resposta atualiza.')
        log("#### ")
        log("-" * 100)
        log("")
        ret = 99
        return(False,False,False,False,False,False)

    ufi      = sys.argv[1].upper()
    mesanoi  = sys.argv[2].upper()
    mesi     = sys.argv[2][:2].upper()
    anoi     = sys.argv[2][2:].upper()

  
    return(ufi,mesanoi,mesi,anoi,iei,invai)

    
if __name__ == "__main__":
    global ret
    
    log('#'*100)
    log("# ")  
    log("# - INICIO - REGISTRO 1400 - MÓDULO CHAMADOR")
    log("# ")
    log('#'*100)
    ret = 0
    retorno = parametros()
      
    ufi = retorno[0]
    if (retorno[0] != False):
        ret      = 0
        ufi      = retorno[0]
        mesanoi  = retorno[1]
        mesi     = retorno[2]
        anoi     = retorno[3]
        iei      = retorno[4]
        invai    = retorno[5]

        log("-"*100)
        log("# Processando ENXERTO SPED para os seguintes parâmetros:")
        log("#    UF      = ",ufi)
        log("#    MÊS     = ",mesi)
        log("#    ANO     = ",anoi)
        log("#    IE      = ",iei)
        log("#  INVA      = ",invai)
        log("-"*100)
        
        if(ufi == 'RJ'):
            log("... Chamando o REGISTRO 1400 DO RJ ...")
            import registro_1400_rj
            ret = registro_1400_rj.main()
        elif(ufi == 'SP'):
            log("... Chamando o REGISTRO 1400 DE SP ...")
            import registro_1400_sp
            ret = registro_1400_sp.main()            
        else:
            log("ERRO - Não foi encontrado o script para o estado informado: ",ufi)
            ret = 99
        
    log('#'*100)
    log("# ")  
    log("# - FIM - REGISTRO 1400 - MÓDULO CHAMADOR")
    log("# ")
    log("#"*100)
    if(ret != 0):
        log("ERRO - VERIFIQUE AS MENSAGENS ANTERIORES PARA IDENTIFICAR O ERRO. ",ret)
    log("Codigo de saida = ",ret)
    sys.exit(ret)
    


    


