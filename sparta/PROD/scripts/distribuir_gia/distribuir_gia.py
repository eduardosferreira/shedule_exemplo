#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: prepara_arquivos_GIA.py
  CRIACAO ..: 27/10/2020
  AUTOR ....: Airton Borges da Silva Filho / KYROS Consultoria
  DESCRICAO : Renomeia os arquivos REGERADOS E PROTOCOLADOS nas pastas /arquivos/GIA<UF>/REGERADOS e
                /arquivos/GIA<UF>/REGERADOS para os nomes padronizados 
                e move para as sub-pastas necessárias para o processamento.
                <UF> é parâmetro informado no comando.
----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 04/02/2020 - Airton Borges da Silva Filho / KYROS Consultoria - Criacao do script.
    * 03/03/2021 - Airton Borges da Silva Filho / KYROS Consultoria - Alterado o padrão do nome dos arquivos de AAAAMM para MMAAAA
    * 18/08/2021 - Marcelo Gremonesi            / Kyros Consultoria - Adquacoes para o novo painel 
    * 22/02/2022 - Eduardo da Silva Ferreira - Kyros Tecnologia
                 - [PTITES-1633] Padrão de diretórios do SPARTA

----------------------------------------------------------------------------------------------
"""

import sys
import os
dir_base = os.path.join( os.path.realpath('.').split('/PROD/')[0], 'PROD') if os.path.realpath('.').__contains__('/PROD/') else os.path.join( os.path.realpath('.').split('/DEV/')[0], 'DEV')
sys.path.append(dir_base)
import configuracoes


import datetime
import sys
import shutil
import glob
import unicodedata
from pathlib import Path
from os import rename
from os import listdir
global ret

import comum
import sql


separadorDiretorio = ('/' if os.name == 'posix' else '\\')
SD=separadorDiretorio
disco = ('' if os.name == 'posix' else 'D:')

def non_ascii_to_ascii(string: str) -> str:
    ascii_only = unicodedata.normalize('NFKD', string)\
        .encode('ascii', 'ignore')\
        .decode('ascii')
    return ascii_only

def encodingDoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r', encoding='iso-8859-1')
        fd.read()
        fd.close()
    except :
        return 'utf-8'

    return 'iso-8859-1'

def tail(f, n):
    assert n >= 0
    pos, lines = n+1, []
    while len(lines) <= n:
        try:
            f.seek(-pos, 2)
        except IOError:
            f.seek(0)
            break
        finally:
            lines = list(f)
        pos *= 2
    return lines[-n:]

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
            log("# ", qdade ," - ",f )
    else: 
        nomearq=""
    return(nomearq)

def normaliza(mascara,diretorio):
    data_criacao = lambda f: f.stat().st_ctime
    data_modificacao = lambda f: f.stat().st_mtime
    qdade = 0
    nomearq = [] 
    directory = Path(diretorio)
    files = directory.glob(mascara)
    sorted_files = sorted(files, key=data_modificacao, reverse=False)
    if sorted_files:
        for f in sorted_files:
            old=str(f)
            new=non_ascii_to_ascii(old)
            rename(old, new)
    return

def lista_arquivos(mascara,diretorio):
    data_criacao = lambda f: f.stat().st_ctime
    data_modificacao = lambda f: f.stat().st_mtime
    qdade = 0
    nomearq = [] 
    directory = Path(diretorio)
    files = directory.glob(mascara)
    sorted_files = sorted(files, key=data_modificacao, reverse=False)
    if sorted_files:
        for f in sorted_files:
            if os.path.isfile(f):    
                qdade = qdade + 1
                nomearq.append(f)
                log("# ",qdade ," - ",f )
    else: 
        nomearq = []
    return(nomearq)

def finalok(f):
    arqok = False
    nlin=1
    try :
        fd = open(f,'r') 
        lin = fd.readline()
    except :
        fd = open(f,'r', encoding=encodingDoArquivo(f))
        lin = fd.readline() 
    try :    
        while (lin):
            nlin = nlin + 1
            if (lin.startswith('|9999|')):
                arqok = True
                break
            else:
                lin = fd.readline()
    except Exception as e :
        arqok = False 
        log("#")
        log("#### ERRO. - DADO ILEGÍVEL ENCONTRADO NA LINHA ", '{:,}'.format(nlin).replace(',','.'))
        log("#### ERRO. - CÓDIGO INTERNO DO ERRO NO SYSTEMA = ", e)
        log("#### ERRO. - PROCURA PELO |9999| INTERROMPIDA DEVIDO A ESTE ERRO. ")
        
        if (str(e).startswith("'charmap'")):
            ec=str(e).split(' ')[5]
            log("#### ERRO. - CÓDIGO DO CARACTERE NÃO RECONHECIDO = ", ec)
    fd.close()
    return(arqok)

def retornaIDArquivo(path) :
    try :
        fd = open(path,'r') 
        lin = fd.readline()
    except :
        fd = open(path,'r', encoding=encodingDoArquivo(path))
        lin = fd.readline()
    fd.close()
    if lin and lin.startswith('|0000|') :
        ano = lin.split('|')[4][4:] 
        mes = lin.split('|')[4][2:4]
        uf = lin.split('|')[9]
        insc = lin.split('|')[10]
        compet_i = lin.split('|')[4]
        compet_f = lin.split('|')[5]
        return [uf, insc, compet_i, compet_f, mes, ano] or [ False, False, False, False, False, False]
    return False, False, False, False, False, False


def retornaIDArquivoGIA(path) :
    try :
        fd = open(path,'r') 
        lin = fd.readline()
    except :
        fd = open(path,'r', encoding=encodingDoArquivo(path))
        lin = fd.readline()
    if lin and lin.startswith('0101') :
        amd_g = lin[4:12] 
        hms_g = lin[12:18]
        lin = fd.readline()
        if lin and lin.startswith('05') :
            ie = lin[2:14] 
            am_d = lin[37:43]    
            return [ie, am_d, amd_g, hms_g] or [ False, False, False, False]      
    return False, False, False, False

def validauf(uf):
    return(True if (uf.upper() in ('AC','AL','AM','AP','BA','CE','DF','ES','GO','MA','MG','MS','MT','PA','PB','PE','PI','PR','RJ','RN','RO','RR','RS','SC','SE','SP','TO')) else False)
          
def dtf():
    return (datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'))

def processar():
    global variaveis
    ufi = ""
     
#### Recebe, verifica e formata os argumentos de entrada.  
    if (len(sys.argv) == 2 and validauf(sys.argv[1].upper()) ): 
        ufi = sys.argv[1].upper()
    else :
        log("-" * 110)
        log("#### ")
        log('#### ERRO - Erro nos parametros do script.')
        log("#### ")
        log('#### Exemplo de como deve ser :')
        log('####      %s <UF> '%(sys.argv[0]))
        log("#### ")
        log('#### Onde')
        log('####      <UF>     = estado. Ex: SP')
        log("#### ")
        log('#### Portanto, se o estado = SP, o comando correto deve ser :')  
        log('####      %s SP '%(sys.argv[0]))  
        log("#### ")
        log('#### ')
        log("-" * 110)
        log("")
        log("Retorno = 99") 
        ret = 99
        return(99)
    
#### Monta caminho e nome do destino
   
    ret = 0
    dir_arquivos = str(SD) + 'arquivos' + str(SD) + 'GIA' + ufi + SD # # [PTITES-1633] # disco + configuracoes.diretorio_arquivos + ufi + SD
    dir_protocolados = os.path.join(dir_arquivos , 'PROTOCOLADOS')
    dir_regerados = os.path.join(dir_arquivos, 'REGERADOS') 
    dir_nova_base_gia = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'GIA') # [PTITES-1633]
    dir_nova_base_gia_entrada = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'GIA') # [PTITES-1633]
    dir_nova_protocolados = os.path.join(dir_nova_base_gia_entrada, 'PROTOCOLADOS', ufi) # [PTITES-1633]
    dir_nova_regerados = os.path.join(dir_nova_base_gia, 'REGERADOS', ufi) # [PTITES-1633]
    # [PTITES-1633]
    log('DIRETORIO DESTINO.:',dir_nova_base_gia, dir_nova_base_gia_entrada, dir_nova_protocolados, dir_nova_regerados)  
    if not os.path.isdir(dir_nova_base_gia) :
        os.makedirs(dir_nova_base_gia)
    if not os.path.isdir(dir_nova_base_gia_entrada) :
        os.makedirs(dir_nova_base_gia_entrada)
    if not os.path.isdir(dir_nova_protocolados) :
        os.makedirs(dir_nova_protocolados)
    if not os.path.isdir(dir_nova_regerados) :
        os.makedirs(dir_nova_regerados)
    # [PTITES-1633]
     
    log('DIRETORIO DOS ARQUIVOS.:',dir_arquivos)  
    
    if not os.path.isdir(dir_protocolados) :
        os.makedirs(dir_protocolados)
    if not os.path.isdir(dir_regerados) :
        os.makedirs(dir_regerados)

    mascara_regerado = "*.TXT"
    mascara_protocolado = "*"
        
    log("-"* 110)    
    log("#")
    log("# Lista de arquivos REGERADOS a serem processados ...")    
    
    
#    arquivos = listdir('.')
#    for arquivo in arquivos:
#        rename(arquivo, arquivo.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").replace("ã","a").replace("ç","c").replace(" ","_").replace(",","").replace("õ","o"))
#        print arquivo.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").replace("ã","a").replace("ç","c").replace(" ","_").replace(",","").replace("õ","o");

    # [PTITES-1633]
    log('Diretorio.:', dir_arquivos, dir_nova_regerados)
    normaliza("*.*", dir_arquivos)
    normaliza("*.*", dir_nova_regerados)
    listaregerados = []
    lista = lista_arquivos(mascara_regerado,dir_arquivos)
    if lista:
       listaregerados.extend(lista)     
    lista = lista_arquivos(mascara_regerado,dir_nova_regerados)
    if lista:
       listaregerados.extend(lista)     
    # [PTITES-1633]
    log('LISTA ARQ REGERADO' ,listaregerados)
    log("#")
    log("-"* 110)
    log("#")
    log("# Lista de arquivos PROTOCOLADOS a serem processados ...")   
    # [PTITES-1633]
    log('Diretorio.:', dir_protocolados, dir_nova_protocolados)
    normaliza("*.*",dir_arquivos)
    listaprotocolados = []
    lista = lista_arquivos(mascara_protocolado,dir_protocolados)
    if lista:
       listaprotocolados.extend(lista)
    lista = lista_arquivos(mascara_protocolado,dir_nova_protocolados)
    if lista:
       listaprotocolados.extend(lista)
    # [PTITES-1633]
    #log(listaprotocolados)
    log("#")
    log("-"* 110)   
    log("#")    

    ####  REGERADOS

    log("###DEBUG###, listaregerados    = ",listaregerados )
    log("###DEBUG###, listaprotocolados = ",listaprotocolados )

    log("# ")
    log("# Organizando os arquivos REGERADOS...")
    log("#")
    for arquivo in listaregerados:
        log("-"*110)
        iea, am_d, amd_g, hms_g = retornaIDArquivoGIA(arquivo)
        
     
        if (iea and am_d and amd_g and hms_g):
            aa = am_d[:4]
            ma = am_d[4:]
            nomd = am_d[4:]+ am_d[0:4]
            
            nova_pasta = os.path.join(dir_nova_regerados,aa,ma)  # [PTITES-1633] os.path.join(dir_regerados,aa,ma)
            mascara = "GIA"+ ufi + "_" + nomd + "_" + iea + "_REG_V*.txt"
            
            
            log("# ")
            log("# Arquivos existentes que possuem o nome padrão...")
            ultimo_arquivo = nome_arquivo(mascara,nova_pasta)
            proximo = "001"
            if (ultimo_arquivo == ""):
                novo_nome = "GIA"+ ufi + "_" + nomd + "_" + iea + "_REG_V001.txt"
            else:
                proximo = '{:03d}'.format(int((str(ultimo_arquivo).split(".")[0]).split("_")[4][1:]) + 1)
                novo_nome = "GIA"+ ufi + "_" + nomd + "_" + iea + "_REG_V"+proximo+".txt"
            novo_arquivo = os.path.join(nova_pasta,novo_nome)
            velho_arquivo = os.path.join(dir_regerados,arquivo)
            if (not os.path.isfile(velho_arquivo)): # [PTITES-1633]
                velho_arquivo = os.path.join(dir_nova_regerados,arquivo) # [PTITES-1633]
            
            bad_nome_novo  = os.path.join(dir_nova_regerados ,"GIA"+ ufi + "_" + nomd + "_" + iea + "_REG_V" + proximo + ".BAD") # [PTITES-1633]
            bad_nome_velho = str(arquivo).split(".")[0] + ".BAD"
            if (os.path.isfile(bad_nome_velho)):# [PTITES-1633]
                bad_nome_velho = os.path.join(dir_nova_regerados,(str(arquivo).split(".")[0] + ".BAD")) # [PTITES-1633]
            
            log("#")
            log("# Definições de padronização e distribuição :")
            log("# Arquivo origem ...............: ",velho_arquivo)
            log("# Sera movido e renomeado para .: ",novo_arquivo)
            log("#")
            
            if not os.path.isdir(nova_pasta) :
                os.makedirs(nova_pasta)
 
            if (os.path.isfile(bad_nome_velho)):
                shutil.move(bad_nome_velho,bad_nome_novo) # [PTITES-1633], copy_function = shutil.copytree)

            if (not os.path.isfile(novo_arquivo)):
                shutil.move(velho_arquivo,novo_arquivo) # [PTITES-1633], copy_function = shutil.copytree)
            else: 
                log("#"*110)
                log("#")
                log("#### ERRO. - ARQUIVO JÁ EXISTE NA PASTA DESTINO.")
                log("#### ERRO. - REMOVA O ARQUIVO ORIGEM OU O ARQUIVO DESTINO")
                log("#### ERRO. - ARQUIVO ORIGEM: ",velho_arquivo)
                log("#### ERRO. - ARQUIVO DESTINO: ",novo_arquivo)
                log("#")
                log("#"*110)
                ret = 99
        else:
            log("#"*110)
            log("#")
            log("#### ERRO. - ARQUIVO INVÁLIDO.")
            log("#### ERRO. - REMOVA OU CORRIJA O ARQUIVO ", arquivo)
            log("#")
            log("#"*110)
            ret = 99
    log("-"*110)
    

    ####  PROTOCOLADOS


    log("# ")
    log("# Organizando os arquivos PROTOCOLADOS...")
    log("#")
    
   
    for arquivo in listaprotocolados:
        log("-"*110)


        iea, am_d, amd_g, hms_g = retornaIDArquivoGIA(arquivo)
        
        if (iea and am_d and amd_g and hms_g):
            aa = am_d[:4]
            ma = am_d[4:]
            nomd = am_d[4:]+ am_d[0:4]
            nova_pasta = os.path.join(dir_nova_protocolados,aa,ma)  # [PTITES-1633] # os.path.join(dir_protocolados,aa,ma)  
            mascara = "GIA"+ ufi + "_" + nomd + "_" + iea + "_PROT_V*.txt"
            
            log("# ")
            log("# Arquivos existentes que possuem o nome padrão...")
            ultimo_arquivo = nome_arquivo(mascara,nova_pasta)
            proximo = "001"
            if (ultimo_arquivo == ""):
                novo_nome = "GIA"+ ufi + "_" + nomd + "_" + iea + "_PROT_V001.txt"
            else:
                proximo = '{:03d}'.format(int((str(ultimo_arquivo).split(".")[0]).split("_")[4][1:]) + 1)
                novo_nome = "GIA"+ ufi + "_" + nomd + "_" + iea + "_PROT_V"+proximo+".txt"
            novo_arquivo = os.path.join(nova_pasta,novo_nome)
            velho_arquivo = os.path.join(dir_protocolados,arquivo)
            if (not os.path.isfile(velho_arquivo)): # [PTITES-1633]
                velho_arquivo = os.path.join(dir_nova_protocolados,arquivo) # [PTITES-1633]
                
            log("#")
            log("# Definições de padronização e distribuição :")
            log("# Arquivo origem ...............:  ",velho_arquivo)
            log("# Sera movido e renomeado para .: ",novo_arquivo)
            log("#")
 
            if not os.path.isdir(nova_pasta) :
                os.makedirs(nova_pasta)
            if (not os.path.isfile(novo_arquivo)):
                shutil.move(velho_arquivo,novo_arquivo) # [PTITES-1633], copy_function = shutil.copytree)
            else: 
                log("#"*110)
                log("#")
                log("#### ERRO. - ARQUIVO JÁ EXISTE NA PASTA DESTINO.")
                log("#### ERRO. - REMOVA O ARQUIVO ORIGEM OU O ARQUIVO DESTINO")
                log("#### ERRO. - ARQUIVO ORIGEM: ",velho_arquivo)
                log("#### ERRO. - ARQUIVO DESTINO: ",novo_arquivo)
                log("#")
                log("#"*110)
                ret = 99
        else:
            log("#"*110)
            log("#")
            log("#### ERRO. - INVÁLIDO.")
            log("#### ERRO. - REMOVA OU CORRIJA O ARQUIVO ", arquivo)
            log("#")
            log("#"*110)
            ret = 99
    log("-"*110)
 
    return(ret) 

if __name__ == "__main__":
    log('#'*110)
    log("# ")  
    log("# ",dtf() , " - INICIO - PREPARA ARQUIVOS - DISTRIBUIR GIA")
    log("# ")
    log('#'*110)
    variaveis = comum.carregaConfiguracoes(configuracoes)
    ret = processar()
       
    log('#'*110)
    log("# ")  
    log("# ",dtf() , " - FIM - PREPARA ARQUIVOS GIA")
    log("# ")
    log("#"*110)

    log("Codigo de saida = ",ret)
    sys.exit(ret)




