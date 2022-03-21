#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-

"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: Enxerto_SPED_RJ
  CRIACAO ..: 15/03/2021
  AUTOR ....: Airton Borges da Silva Filho / KYROS Consultoria
  DESCRICAO :

            Este script recebe todos os parametros necessarios a execucao dos enxertos de cada estado e chama o script especifico.

            VARIÁVEIS:
                <UF>     =  Estado (Unidade Federativa) ex: RJ
                <MMAAAA> =  MM: mes, numérico com dois dígitos. AAAA: ano, numérico com quatro dígitos.  ex: 012021
                <IE>     =  Inscriçao estadual. numérico.  ex: 77452443
                <1400SP: S/N>    =  A letra S ou a letra N.  ex: S . Somente para SP, enxertar o bloco 1400 ou nao
                <1600SP: S/N>    =  A letra S ou a letra N.  ex: S . Somente para SP, enxertar o bloco 1600 ou nao
                <1600RJ: S/N>    =  A letra S ou a letra N.  ex: S . Somente para RJ. S para forcar o enxerto do bloco 1600 sempre do REGERADO
                                                                                      N para enxertar o bloco 1600 de acordo com a data dos dados.
                        Detalhes 1600RJ: Somente para o RJ.                                                                                     
                                S:  Faz o enxerto a partir do REGERADO independente da data dos dados.
                                N:  Faz o enxerto a partir do REGERADO se a data dos dados for menor ou igual a 072017 ou
                                    Faz o enxerto a partir do PROTOCOLADO se a data dos dados for maior ou igual a 082017
                                    
            OBS:
                Caso informados dados para um estano que não os necessite, eles serão despresados na chamada do script específico.
                Caso falte dados para o estado informado, uma mensagem de erro será apresentada e o script específico do estado não será chamado.
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
  
    if (len(sys.argv) == 8        
        and validauf(sys.argv[1])
        and len(sys.argv[2])==6  
        and int(sys.argv[2][:2])>0 
        and int(sys.argv[2][:2])<13
        and int(sys.argv[2][2:])<=datetime.datetime.now().year
        and int(sys.argv[2][2:])>2014
        and (sys.argv[4].upper() in ("S", "'S'", '"S"',"N", "'N'", '"N"'))
        and (sys.argv[5].upper() in ("S", "'S'", '"S"',"N", "'N'", '"N"'))
        and (sys.argv[6].upper() in ("S", "'S'", '"S"',"N", "'N'", '"N"'))
        ):
        
        iei=sys.argv[3].upper()
        iei = re.sub('[^0-9]','',iei)
        if ( (iei == "") or (iei == "''") or (iei == '""') or (int("0"+iei) == 0)):
            iei = "*"
             
    else :
        log("-" * 100)
        log("#### ")
        log('#### ERRO - Erro nos parametros do script.')
        log("#### ")
        log('#### Exemplo de como deve ser :')
        log('####      %s <UF> <MMAAAA> <IE> <S|N> <S|N> <S|N> <SERIES>'%(sys.argv[0]))
        log("#### ")
        log('#### Onde')
        log('####      <UF>     = mês e ano. Ex: Para junho de 2020 informe 062020')
        log('####      <MMAAAA> = mês e ano. Ex: Para junho de 2020 informe 062020')
        log('####      <IE>     = Inscição Estadual.')
        log('####      <S|N>    = Para SP, Enxerta o bloco 1400 ou não')
        log('####      <S|N>    = Para SP, Enxerta o bloco 1600 do REGERADO independente de data ou N para enxertar do')
        log('####                 PROTOCOLADO se data >= 08/2017')
        log('####      <S|N>    = Para o RJ, Se informado "S", realiza enxerto do bloco 1600 do REGERADO independente de data.')
        log('####                 Para o RJ, Se informado "N", realiza enxerto do bloco 1600 do REGERADO até 07/2017 ou do')
        log('####                 PROTOCOLADO para datas maiores ou iguais a 08/2017.')
        log('####      <SERIES> = Series a serem substituidas nos blocos D695, D696, D697 .')
        log('####                 caso não informado, todas as series serão substituidas.')
        log("#### ")
        log("-" * 100)
        log("")
        ret = 99
        return(False,False,False,False,False,False,False,False)

    ufi      = sys.argv[1].upper()
    mesanoi  = sys.argv[2].upper()
    mesi     = sys.argv[2][:2].upper()
    anoi     = sys.argv[2][2:].upper()
    sp1400i  = sys.argv[4].upper()
    sp1600i  = sys.argv[5].upper()
    rj1600i  = sys.argv[6].upper()
    series   = sys.argv[7].upper()
  
    return(ufi,mesanoi,mesi,anoi,iei,sp1400i,sp1600i,rj1600i,series)

    
if __name__ == "__main__":
    global ret
    
    log('#'*100)
    log("# ")  
    log("# - INICIO - ENXERTO_SPED MÓDULO CHAMADOR")
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
        sp1400i  = retorno[5]
        sp1600i  = retorno[6]
        rj1600i  = retorno[7]
        series   = retorno[8]

        log("-"*100)
        log("# Processando ENXERTO SPED para os seguintes parâmetros:")
        log("#    UF      = ",ufi)
        log("#    MÊS     = ",mesi)
        log("#    ANO     = ",anoi)
        log("#    IE      = ",iei)
        log("#    SP1400  = ",sp1400i)
        log("#    SP1600  = ",sp1600i)
        log("#    RJ1600  = ",rj1600i)
        log("-"*100)

        comando = ""
        
        if(ufi == 'RJ'):
            log("Chamando o enxerto de RJ")
            comando = "enxerto_sped_rj.py"
            param = mesanoi + " " + iei + " " + rj1600i  
           
        elif(ufi == 'SP'):
            log("Chamando o enxerto de SP")
            comando = "enxerto_sped_sp.py"
            param = mesanoi  + " " + sp1400i + " " + sp1600i + " " + iei + ' "%s"'%(series)
            
        elif(ufi == 'PE'):
            log("Chamando o enxerto do PE")
            comando = "enxerto_sped_pe.py" 
            param = ufi + " " + mesanoi + " " + iei 
            
        else:
            log("ERRO - Não foi encontrado o script para o estado informado: ",ufi)
            ret = 99

        if ( comando != "" ):
            log("Comando    = ", comando)
            log("parametros = ",param)
            
            compar = "./"+ comando +" "+ param
           
            processo = subprocess.Popen(compar,stdout=subprocess.PIPE,stdin=subprocess.PIPE, stderr=subprocess.PIPE,shell=True )           
  
            while processo.poll() is None:
                time.sleep(5)
            
            out, err = processo.communicate() 
            print('L O G    D A    E X E C U C A O'.center(100,'-'))
            for lin in out.decode('utf-8').strip('\n').split('\t'):
                print('>', lin)
            print('F I M    D O    L O G    D A    E X E C U C A O'.center(100,'-'))
            
            ret     = processo.returncode

        
    log('#'*100)
    log("# ")  
    log("# - FIM - ENXERTO_SPED MÓDULO CHAMADOR")
    log("# ")
    log("#"*100)
    if(ret != 0):
        log("ERRO - VERIFIQUE AS MENSAGENS ANTERIORES PARA IDENTIFICAR O ERRO. ",ret)
    log("Codigo de saida = ",ret)
    sys.exit(ret)
    


    
