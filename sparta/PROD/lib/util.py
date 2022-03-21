"""
----------------------------------------------------------------------------------------------
  BIBLIOTECA .: sql.py
  CRIACAO ....: 27/03/2021
  AUTOR ......: VICTOR SANTOS CARDOSO / KYROS TECNOLOGIA
  DESCRICAO ..: Biblioteca de funcoes pontuais usadas pelos scripts.

----------------------------------------------------------------------------------------------
  HISTORICO ..: 
    
----------------------------------------------------------------------------------------------
"""

import datetime
import cx_Oracle
import os
import sys
import atexit
import traceback
import re
from pathlib import Path

def validaano(ano):
    return(True if (len(ano) == 4 and int(ano) > 2000 and int(ano) <= (datetime.datetime.now().year )) else False)

def validames(mes):
    return(True if (len(mes) == 2 and int(mes) > 0 and int(mes) < 13) else False)

def dtf():
    return (datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'))


def validauf(uf):
    return(True if (uf.upper() in ('AC','AL','AM','AP','BA','CE','DF','ES','GO','MA','MG','MS','MT','PA','PB','PE','PI','PR','RJ','RN','RO','RR','RS','SC','SE','SP','TO')) else False)

def valida_ano(ano):
    return(ano if (len(ano) == 4 and int(ano) > 2005 and int(ano) <= (datetime.datetime.now().year )) else "#")


def valida_mes(mes):
    return(mes if (len(mes) == 2 and int(mes) > 0 and int(mes) < 13) else "#")


def valida_ie(ie):
    ie = re.sub('[^0-9]','',ie)
    return( "#" if ( (ie == "") or (ie == "''") or (ie == '""') or (int("0"+ie) == 0)) else ie )

def valida_uf(uf):
    return(uf.upper() if(validauf(uf)) else "#")

def nome_arquivo(mascara,diretorio):
    qdade = 0
    nomearq = "" 
    directory = Path(diretorio)
    files = directory.glob(mascara)
    sorted_files = sorted(files, reverse=False)
    if sorted_files:
        for f in sorted_files:
            qdade = qdade + 1
            nomearq = f
    else: 
        log('ERRO : Arquivo %s nÃ£o estÃ¡ na pasta %s'%(mascara,diretorio))
        log("")
    return(nomearq)

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
            ie = str(f).split("_")[4]
            log("   ",qdade, " => ", f)
            try:
                ies.index(str(f).split("_")[4])
            except:
                ies.append(str(f).split("_")[4])
                continue
            
    else: 
        log('ERRO : Arquivo %s não está na pasta %s'%(mascara,diretorio))
        log("")
        ret=99
        return("")
    log(" ")
    return(ies)

def lista_ies_existentes(mascara,diretorio):
    
    qdade = 0
    ies = []
    directory = Path(diretorio)
    files = directory.glob(mascara)
    sorted_files = sorted(files, reverse=False)
    if sorted_files:
        log("Arquivos encontrados: ")
        for f in sorted_files:
            qdade = qdade + 1
            ie = str(f).split("_")[-3]
            log("   ",qdade, " => ", f, " => ", str(ie))
            try:
                ies.index(str(f).split("_")[-3])
            except:
                ies.append(str(f).split("_")[-3])
                continue
            
    else: 
        log('ERRO : Arquivo %s não está na pasta %s'%(mascara,diretorio))
        log("")

    log(" ")
    return(ies)

def isCpfValid(cpf):
    """ If cpf in the Brazilian format is valid, it returns True, otherwise, it returns False. """
    if not isinstance(cpf,str):
        return False
    cpf = re.sub("[^0-9]",'',cpf)
    if cpf=='00000000000' or cpf=='22222222222' or cpf=='33333333333' or cpf=='44444444444' or cpf=='55555555555' or cpf=='66666666666' or cpf=='77777777777' or cpf=='88888888888' or cpf=='99999999999':
        return False
    if len(cpf) != 11:
        return False
    sum = 0
    weight = 10
    """ Calculating the first cpf check digit. """
    for n in range(9):
        sum = sum + int(cpf[n]) * weight
        weight = weight - 1
    verifyingDigit = 11 -  sum % 11
    if verifyingDigit > 9 :
        firstVerifyingDigit = 0
    else:
        firstVerifyingDigit = verifyingDigit
    """ Calculating the second check digit of cpf. """
    sum = 0
    weight = 11
    for n in range(10):
        sum = sum + int(cpf[n]) * weight
        weight = weight - 1
    verifyingDigit = 11 -  sum % 11
    if verifyingDigit > 9 :
        secondVerifyingDigit = 0
    else:
        secondVerifyingDigit = verifyingDigit
    if cpf[-2:] == "%s%s" % (firstVerifyingDigit,secondVerifyingDigit):
        return True
    return False

def isCnpjValid(cnpj):
    """ If cnpf in the Brazilian format is valid, it returns True, otherwise, it returns False. """

    # Check if type is str
    if not isinstance(cnpj,str):
        return False

    cpf = re.sub("[^0-9]",'',cnpj)

    if len(cpf) != 14:
        return False

    if cnpj == "00011111111111":
        return True

    sum = 0
    weight = [5,4,3,2,9,8,7,6,5,4,3,2]

    """ Calculating the first cpf check digit. """
    for n in range(12):
        value =  int(cpf[n]) * weight[n]
        sum = sum + value


    verifyingDigit = sum % 11

    if verifyingDigit < 2 :
        firstVerifyingDigit = 0
    else:
        firstVerifyingDigit = 11 - verifyingDigit

    """ Calculating the second check digit of cpf. """
    sum = 0
    weight = [6,5,4,3,2,9,8,7,6,5,4,3,2]
    for n in range(13):
        sum = sum + int(cpf[n]) * weight[n]

    verifyingDigit = sum % 11

    if verifyingDigit < 2 :
        secondVerifyingDigit = 0
    else:
        secondVerifyingDigit = 11 - verifyingDigit

    if cpf[-2:] == "%s%s" % (firstVerifyingDigit,secondVerifyingDigit):
        return True
    return False