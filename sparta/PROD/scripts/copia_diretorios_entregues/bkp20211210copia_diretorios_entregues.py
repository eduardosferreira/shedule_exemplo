#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: copia_diretorios_entregues.py
  CRIACAO ..: 10/12/2021
  AUTOR ....: Airton Borges da Silva Filho / KYROS Consultoria
  DESCRICAO :
      Parâmetros Input: Filial, Série (opcional), Ano Mes Inicio, Ano Mes Fim
      compara os arquivos nas pastas:
      


    ./copia_diretorios_entregues.py ie filial ddmmyyyyinicial ddmmyyyyfinal serie
    ./copia_diretorios_entregues.py "108383949112" "0001" "012015" "012015" ""
    ./copia_diretorios_entregues.py "108383949112" "0001" "012017" "012017" "170000397"
    
    
    
"""
import sys
import os


SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)

import configuracoes
import comum
import datetime 
from os import listdir
import pathlib
from os.path import isfile, join

comum.carregaConfiguracoes(configuracoes)
ret = 0 
log = comum.log
status_final = 0
log.gerar_log_em_arquivo = True

ret                 = 0
CC_IE               = ""
CC_FILIAL           = ""
DT_MESANO_INICIO    = ""
DT_MESANO_FIM       = ""
CC_SERIE            = ""



def gravar(dest,dados):
    if(os.path.isfile(dest)):
        arq_csv = open(dest,"a")
    else:
        arq_csv = open(dest,"w")
        arq_csv.write("ULTIMA_ENTREGA;REGERADO;TIPO_ARQUIVO;VOLUME;RESULTADO_COMPARACAO;DESCRICAO_COMPARACAO;CONTEUDO_COMPARACAO;DATA_HORA_COMPARACAO\n")
    arq_csv.write(dados)
    arq_csv.close()
    
    
def anomesEntre(DT_MESANO_INICIO,DT_MESANO_FIM):
    
    print("DT_MESANO_INICIO = ",DT_MESANO_INICIO)
    print("DT_MESANO_FIM = ",DT_MESANO_FIM)
    input("Continua ?") 
    
     
    lst_mesano = []
    
    mesini =  str(DT_MESANO_INICIO[3:5].zfill(2))
    anoini =  str(DT_MESANO_INICIO[6:].zfill(4)) 
    anomesini = anoini+mesini
    
    mesfim =  str(DT_MESANO_FIM[3:5].zfill(2))
    anofim =  str(DT_MESANO_FIM[6:].zfill(4))
    anomesfim = anofim+mesfim
  
    anomesatu = anomesini
    anoatu = anoini
    mesatu = mesini
   
    print("anomesini = ", anomesini)
    print("anomesfim = ", anomesfim)    
    print("anomesatu = ", anomesatu)    
    print("")
    
    while anomesfim >= anomesatu:
        lst_mesano.append(anomesatu[2:])
        
        if (int(mesatu) == 12):
            mesatu = '01'
            anoatu = str(int(anoatu)+1).zfill(4)
        else:
            mesatu = str(int(mesatu) + 1).zfill(2)
            
        anomesatu = str(anoatu).zfill(4) + str(mesatu).zfill(2)        
        
        print("anomesatu = ", anomesatu)    
    
    print("")        
    print(lst_mesano)
    input("Continua ?")    
        
    return(lst_mesano)
    
    
    


def selectPastas(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,UFI="SP") :
    log("Gerando lista de pastas existentes a copiar....")

    pasta = configuracoes.dir_base_copia

    lista_datas = anomesEntre(DT_MESANO_INICIO,DT_MESANO_FIM)
    
    lst_serie = CC_SERIE.split(",")
    
    for serie in lst_serie:
        print( "serie = ", serie)
    
    
    
    lst_path_regerado = []    
    for dat in lista_datas:
        mask = os.path.join(pasta,UFI,dat,"TBRA",CC_FILIAL,"SERIE")                 
        print(mask)
        
        
        
        
        
        
        
        
        
    
    print("CC_SERIE = ", CC_SERIE)
    
    input("Termina aqui. ")
    
    
    #DEBUG# - para testar em produção:
    #pasta = pasta[68:]
    
    pasta = os.path.join(pasta,UFI,DT_MESANO_INICIO[8:10],DT_MESANO_INICIO[3:5],"TBRA",CC_FILIAL.replace("'",""),"SERIE")
    pasta_inicial = pasta
    
    lst_all_files = []
    lst_series = []

    if not os.path.isdir(pasta):
        #log("ERRO - PASTA:\n", pasta ,"\nnão existe.")
        #ret = 66
        os.mkdir(pasta)
    else:
        if (CC_SERIE =="''" ):
            lst_series = [f for f in listdir(pasta) if not isfile(join(pasta, f))]
        else:
            lst_series.append(CC_SERIE.replace("'",""))
      
        for serie_atu in lst_series:
            pasta = os.path.join(pasta_inicial,serie_atu,PASTA)
            if not os.path.isdir(pasta):
                #log("ERRO - PASTA:\n", pasta ,"\nnão existe.")
                #ret = 66
                os.mkdir(pasta)
            else:
                lst_files = [f for f in listdir(pasta) if isfile(join(pasta, f))]
                for file in lst_files:
                    lst_all_files.append(os.path.join(pasta,file))

    return(lst_all_files)
    
    
    
def NOVOprocessar(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,UFI) :
    ret = 0
    
    reg_obrigacao_tmp = selectArquivos(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,configuracoes.OBRIGACAO,UFI)
    regs_ultima_entrega_tmp = selectArquivos(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,"ULTIMA_ENTREGA",UFI)
    amdhms = str(datetime.datetime.now().strftime('%Y%m%d-%H%M%S'))
    
    #filtra registros que interessam....
    registros=[]
    regs_ultima_entrega=[]
    
    
        
    #monta o nome do arquivo csv com o resultado da comparacao:
    pasta_csv = configuracoes.dir_saida_comp
    nome_csv = "comparativo_" + UFI + "_" + CC_IE + "_" + CC_FILIAL.replace("'","")  + "_" + DT_MESANO_INICIO[6:10] + DT_MESANO_INICIO[3:5] + "_" + CC_SERIE.replace("'","") +"_"+  amdhms +".csv"
    nome_arq_csv = os.path.join(pasta_csv,nome_csv)

    if not os.path.isdir(pasta_csv):
        os.mkdir(pasta_csv)

    print("")        
    log("="*160)
    log("Nome do arquivo com o resultado de todas comparações: ")
    log("   ==>  " + str(nome_arq_csv))
    log("="*160)
    print("")        
  
    arq_csv = open(nome_arq_csv,"w")
    
#    arq_csv.write("ULTIMA_ENTREGA;"+ configuracoes.OBRIGACAO + " - (REGERADO);TIPO_ARQUIVO;VOLUME;RESULTADO_COMPARACAO;DESCRICAO_COMPARACAO;CONTEUDO_ULTIMA_ENTREGA;CONTEUDO "+ configuracoes.OBRIGACAO + " - (REGERADO) \n")
    arq_csv.write(str(configuracoes.OBRIGACAO) + " - (REGERADO);ULTIMA_ENTREGA; TIPO_ARQUIVO;VOLUME;RESULTADO_COMPARACAO;DESCRICAO_COMPARACAO;CONTEUDO_"+str(configuracoes.OBRIGACAO) + " - (REGERADO);CONTEUDO_ULTIMA_ENTREGA \n")



  
    
    #SELECIONA SOMENTE ARQUIVOS TIPOS 'C','D','I' E 'M' EM OBRIGACAO
    for reg in reg_obrigacao_tmp:
        nom = reg.split(".")[-2]
        ext = reg.split(".")[-1]
        tipo = nom[-1]
        if ( (not ext.isdigit()) or (tipo not in ('C','D','I','M'))  ):
            continue
        else:
            registros.append(reg)
   
    qtde = len(registros)
    msg = "Quantidade de arquivos em "+ configuracoes.OBRIGACAO + " = " + str(qtde)
    log(msg)
    #arq_csv.write(";;;;I;" + msg + ";;\n")
    
        
    if ( qtde < 1 ):
        msg = "ERRO - Não existem arquivos em " + configuracoes.OBRIGACAO + " a serem comparados."
        log(msg)
        arq_csv.write(";;;;E;" + msg + ";;\n")
        ret = 55

    



    #SELECIONA SOMENTE ARQUIVOS TIPOS 'C','D','I' E 'M' EM ULTIMA ENTREGA
    for reg in regs_ultima_entrega_tmp:
        nom = reg.split(".")[-2]
        ext = reg.split(".")[-1]
        tipo = nom[-1]
        if ( (not ext.isdigit()) or (tipo not in ('C','D','I','M'))  ):
            continue
        else:
            regs_ultima_entrega.append(reg)
     
    qtdeU = len(regs_ultima_entrega)
    msg = "Quantidade de arquivos em ULTIMA_ENTREGA = " + str(qtdeU)
    log(msg)
    #arq_csv.write(";;;;I;" + msg + ";;\n")
        
    if ( qtdeU < 1 ):
        msg = "ERRO - Não existem arquivos em 'ULTIMA_ENTREGA' a serem comparados."
        log(msg)
        arq_csv.write(";;;;E;" + msg + ";;\n")
        ret = 55
        
   
        
    #DEBUG# = mostra arquivos em obrigacao
    print("")
    print("ARQUIVOS EM 'OBRIGACAO'....:")
    for a in registros:
        print("    - ", a)
        #result_comp =  a + ";;" + tipo + ";" + ext + ";I;; \n"
        #arq_csv.write(result_comp)
    print("-"*160)
        

    a_comparar=[]

        
    #DEBUG# = mostra arquivos em ultima_entrega
    print("")
    print("ARQUIVOS 'ULTIMA_ENTREGA'....:")
    for a in regs_ultima_entrega:
        print("    - ", a)
        #result_comp = ";" + a + ";" + tipo + ";" + ext + ";I;;" + "\n"
        #arq_csv.write(result_comp)
    print("-"*160)

    print("")
    print("")


    #VERIFICA SE TODOS OS ARQUIVOS EM ULTIMA_ENTREGA EXISTEM EM OBRIGACAO
    #arquivo2=ultima_entrega
    log("VERIFICA SE TODOS OS ARQUIVOS EM ULTIMA_ENTREGA EXISTEM EM " + configuracoes.OBRIGACAO )
    for arquivo2 in regs_ultima_entrega:
        
        
        
        
        print("")
        print("")
        log("verificando ",arquivo2)
        
        tipo = arquivo2.split(".")[0][-1]
        ext = arquivo2.split(".")[-1]

        #monta a mascara com o nome do arquivo ULTIMA_ENTREGA. Path e mascara completos.
        arquivo2mask = arquivo2.split(".")[0][:-4] + "???" + arquivo2.split(".")[0][-1] + "." + arquivo2.split(".")[1]
        #pega só o nome do arquivo ultima_entrega
        arquivo_ultima_entrega_temp = arquivo2mask.split("/")[-1]
        #pega só o path de ultima_entrega
        path_ultima_entrega_temp = arquivo2mask.replace(arquivo_ultima_entrega_temp,"")
        #monta só o path de obrigacao.
        path_obrigacao_temp = path_ultima_entrega_temp.replace("ULTIMA_ENTREGA",configuracoes.OBRIGACAO)
        #verifica quantos arquivos de ultima_entrega correspondem aos de obrigacao
        path = list(pathlib.Path(path_obrigacao_temp).glob(arquivo_ultima_entrega_temp))
        
        qo = len(path)
        msg = ""
        if ( qo > 1 ):
            msg = "Existem mais de 1 arquivo correspondente em OBRIGACAO para um em ULTIMA_ENTREGA. Existe mais de um sequencial para o mesmo volume em OBRIGACAO."
            result_comp = ";"+ str(arquivo2) + ";" + tipo + ";" + ext + ";" + "?" + ";" + msg + ";; \n"
            log("ERRO -  ", msg)
            arq_csv.write(result_comp)
            err = 111
        elif ( qo < 1 ):
            msg = "Não existe arquivo correspondente em OBRIGACAO ao arquivo em ULTIMA_ENTREGA."
            result_comp = ";"+ str(arquivo2) + ";" + tipo + ";" + ext + ";" + "?" + ";" + msg + ";; \n"
            log("ERRO -  ",msg)
            arq_csv.write(result_comp)
            err = 111
        elif (qo == 1):
            log("OK - Arquivo ULTIMA_ENTREGA:")
            log("    ====> ",arquivo2)
            log("Existe na pasta OBRIGACAO como:")
            log("    ====> ",path[0])




   #VERIFICA SE TODOS OS ARQUIVOS EM OBRIGACAO EXISTEM EM ULTIMA_ENTREGA
    log("VERIFICA SE TODOS OS ARQUIVOS EM " + configuracoes.OBRIGACAO + " EXISTEM EM ULTIMA_ENTREGA")
    for arquivo1 in registros:
        print("")
        print("")
        log("verificando ",arquivo1)
        
        tipo = arquivo1.split(".")[0][-1]
        ext = arquivo1.split(".")[-1]
        
        #monta a mascara com o nome do arquivo OBRIGACAO. Path e mascara completos.
        arquivo1mask = arquivo1.split(".")[0][:-4] + "???" + arquivo1.split(".")[0][-1] + "." + arquivo1.split(".")[1]
        #pega só o nome do arquivo obrigacao
        arquivo_obrigacao_temp = arquivo1mask.split("/")[-1]
        #pega só o path de obrigacao
        path_obrigacao_temp = arquivo1mask.replace(arquivo_obrigacao_temp,"")
        #monta só o path de ultima_entrega.
        path_ultima_temp = path_obrigacao_temp.replace(configuracoes.OBRIGACAO, "ULTIMA_ENTREGA")
        #verifica quantos arquivos de ultima_entrega correspondem aos de obrigacao
        path = list(pathlib.Path(path_ultima_temp).glob(arquivo_obrigacao_temp))
        qe = len(path)
        msg = ""
        if ( qe > 1 ):
            msg = "Existem mais de 1 arquivo correspondente em ULTIMA_ENTREGA para um em OBRIGACAO. Existe mais de um sequencial para o mesmo volume em ULTIMA_ENTREGA."
            result_comp = str(arquivo1) + ";;" + tipo + ";" + ext + ";?;" + msg + ";; \n"
            log("ERRO -  ", msg)
            arq_csv.write(result_comp)
            err = 111
        elif ( qe < 1 ):
            msg = "Não existe arquivo correspondente em ULTIMA_ENTREGA ao arquivo em OBRIGACAO."
            result_comp = str(arquivo1) + ";;" + tipo + ";" + ext + ";?;" + msg + ";; \n"
            log("ERRO -  ",msg)
            arq_csv.write(result_comp)
            err = 111
        elif (qe == 1):
            log("OK - Arquivo Obrigação:")
            log("    ====> ",arquivo1)
            log("Existe na pasta ULTIMA_ENTREGA como:")
            log("    ====> ",path[0])

            #SELECIONA A DUPLA DE ARQUIVOS PARA COMPARACAO
            a_comparar.append((arquivo1,str(path[0])))
         
            
    print("")
    print("")
    
    
    #COMPARA OS ARQUIVOS SELECIONADOS:
    for comparar in a_comparar:
        arquivo1 = comparar[0]
        arquivo2 = comparar[1]
        err = 0
        print("-"*160)
        log("Comparando os arquivos:")
        log("1 = '",arquivo1,"'")
        log("  com:")
        log("2 = '",arquivo2,"'")
        
       
        if (not os.path.isfile(arquivo1)):
            ret = 11
            continue

        if (not os.path.isfile(arquivo2)):
            ret = 22
            continue

        ##### COMPARA O CONTEUDO DOS ARQUIVOS...
        tipo = arquivo1.split(".")[0][-1]
        ext = arquivo1.split(".")[-1] 

        #transforma os arquivos arq1 e arq2 em listas...
        arq1 = open(arquivo1,"r", encoding=comum.encodingDoArquivo(arquivo1)).read().split("\n")
        arq2 = open(arquivo2,"r", encoding=comum.encodingDoArquivo(arquivo2)).read().split("\n")
        

        if (tipo  in ('C','D','I','M')):
            
            l1=len(arq1)
            l2=len(arq2)
            
            for x in range(0,l1):
                #RETIRAR 10  50 A  59 => CODIGO DO ITEM
                #RETIRAR  6 258 A 263 => ALIQUOTA PIS/PASEP
                #RETIRAR  6 275 A 280 => ALIQUOTA CONFINS
                #RETIRAR 32 300 A 331 => CODIGO AUTENTIFICACAO (331 = FINAL)
                if (tipo == "I"):
                    arq1[x]=arq1[x][0:49]+arq1[x][59:257]+arq1[x][263:274]+arq1[x][280:299]
            for y in range(0,l2):
                #RETIRAR 10  50 A  59 => CODIGO DO ITEM
                #RETIRAR  6 258 A 263 => ALIQUOTA PIS/PASEP
                #RETIRAR  6 275 A 280 => ALIQUOTA CONFINS
                #RETIRAR 32 300 A 331 => CODIGO AUTENTIFICACAO (331 = FINAL)
                if (tipo == "I"):
                    arq2[y]=arq2[y][0:49]+arq2[y][59:257]+arq2[y][263:274]+arq2[y][280:299]
            
            ##### COMPARA A QUANTIDADE DE REGISTROS
            if (l1 != l2):
                log("ERRO - ", arquivo1, " e ", arquivo2, " possuem quantidade de registros diferentes:")
                log("ERRO - ", arquivo1, " = ", len(arq1), " registros")
                log("ERRO - ", arquivo2, " = ", len(arq2), " registros")
                result_comp = arquivo1 + ";"+ arquivo2 + ";" + tipo + ";" + ext + ";" + "#" + ";" + "Quantidade de registros diferentes."+  arquivo1 + "=> " + str(l1) + " e " + arquivo2 + " =>" + str(l2) + ";;" + "\n"
                arq_csv.write(result_comp)
                ret = 55
                err = 1

            else:    
                ##### COMPARA O CONTEUDO DOS ARQUIVOS
                conjunto1 = set(arq1)
                diferencas = conjunto1.difference(arq2)
                ndif = len(diferencas)
                if ( ndif != 0 ):
                    log("ERRO - ", arquivo2, " e ", arquivo1, " possuem ", ndif, " diferença(s).")
                    print("")
                if (ndif > 0):
                    for x in range(0,l1):
                        if (arq1[x] != arq2[x]):
                            result_comp = arquivo1 + ";"+ arquivo2 + ";" + tipo + ";" + ext + ";" + "#" + ";" + "Conteudo diferente" + ";" + arq1[x] + ";" + arq2[x] + " \n"
                            arq_csv.write(result_comp)
                            ret = 55
                            err = 1
                else:
                    #log(arquivo1)
                    #log("e")
                    #log(arquivo2)
                    log("São iguais.")
                    print("")
                    result_comp = arquivo1 + ";"+ arquivo2 + ";" + tipo + ";" + ext + ";" + "=" + ";" + "Arquivos iguais" + ";"+ "''" + "; \n"
                    arq_csv.write(result_comp)

    return ret



















def processar(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,UFI) :
    ret = 0
    reg_obrigacao_tmp = selectArquivos(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,configuracoes.OBRIGACAO,UFI)
    regs_ultima_entrega_tmp = selectArquivos(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,"ULTIMA_ENTREGA",UFI)
    amdhms = str(datetime.datetime.now().strftime('%Y%m%d-%H%M%S'))
    
    #filtra registros que interessam....
    registros=[]
    regs_ultima_entrega=[]
    
    
        
    #monta o nome do arquivo csv com o resultado da comparacao:
    pasta_csv = configuracoes.dir_saida_comp
    nome_csv = "comparativo_" + UFI + "_" + CC_IE + "_" + CC_FILIAL.replace("'","")  + "_" + DT_MESANO_INICIO[6:10] + DT_MESANO_INICIO[3:5] + "_" + CC_SERIE.replace("'","") +"_"+  amdhms +".csv"
    nome_arq_csv = os.path.join(pasta_csv,nome_csv)

    if not os.path.isdir(pasta_csv):
        os.mkdir(pasta_csv)

    print("")        
    log("="*160)
    log("Nome do arquivo com o resultado de todas comparações: ")
    log("   ==>  " + str(nome_arq_csv))
    log("="*160)
    print("")        
  
    arq_csv = open(nome_arq_csv,"w")
    
#    arq_csv.write("ULTIMA_ENTREGA;"+ configuracoes.OBRIGACAO + " - (REGERADO);TIPO_ARQUIVO;VOLUME;RESULTADO_COMPARACAO;DESCRICAO_COMPARACAO;CONTEUDO_ULTIMA_ENTREGA;CONTEUDO "+ configuracoes.OBRIGACAO + " - (REGERADO) \n")
    arq_csv.write(str(configuracoes.OBRIGACAO) + " - (REGERADO);ULTIMA_ENTREGA; TIPO_ARQUIVO;VOLUME;RESULTADO_COMPARACAO;DESCRICAO_COMPARACAO;CONTEUDO_"+str(configuracoes.OBRIGACAO) + " - (REGERADO);CONTEUDO_ULTIMA_ENTREGA \n")



  
    
    #SELECIONA SOMENTE ARQUIVOS TIPOS 'C','D','I' E 'M' EM OBRIGACAO
    for reg in reg_obrigacao_tmp:
        nom = reg.split(".")[-2]
        ext = reg.split(".")[-1]
        tipo = nom[-1]
        if ( (not ext.isdigit()) or (tipo not in ('C','D','I','M'))  ):
            continue
        else:
            registros.append(reg)
   
    qtde = len(registros)
    msg = "Quantidade de arquivos em "+ configuracoes.OBRIGACAO + " = " + str(qtde)
    log(msg)
    #arq_csv.write(";;;;I;" + msg + ";;\n")
    
        
    if ( qtde < 1 ):
        msg = "ERRO - Não existem arquivos em " + configuracoes.OBRIGACAO + " a serem comparados."
        log(msg)
        arq_csv.write(";;;;E;" + msg + ";;\n")
        ret = 55

    



    #SELECIONA SOMENTE ARQUIVOS TIPOS 'C','D','I' E 'M' EM ULTIMA ENTREGA
    for reg in regs_ultima_entrega_tmp:
        nom = reg.split(".")[-2]
        ext = reg.split(".")[-1]
        tipo = nom[-1]
        if ( (not ext.isdigit()) or (tipo not in ('C','D','I','M'))  ):
            continue
        else:
            regs_ultima_entrega.append(reg)
     
    qtdeU = len(regs_ultima_entrega)
    msg = "Quantidade de arquivos em ULTIMA_ENTREGA = " + str(qtdeU)
    log(msg)
    #arq_csv.write(";;;;I;" + msg + ";;\n")
        
    if ( qtdeU < 1 ):
        msg = "ERRO - Não existem arquivos em 'ULTIMA_ENTREGA' a serem comparados."
        log(msg)
        arq_csv.write(";;;;E;" + msg + ";;\n")
        ret = 55
        
   
        
    #DEBUG# = mostra arquivos em obrigacao
    print("")
    print("ARQUIVOS EM 'OBRIGACAO'....:")
    for a in registros:
        print("    - ", a)
        #result_comp =  a + ";;" + tipo + ";" + ext + ";I;; \n"
        #arq_csv.write(result_comp)
    print("-"*160)
        

    a_comparar=[]

        
    #DEBUG# = mostra arquivos em ultima_entrega
    print("")
    print("ARQUIVOS 'ULTIMA_ENTREGA'....:")
    for a in regs_ultima_entrega:
        print("    - ", a)
        #result_comp = ";" + a + ";" + tipo + ";" + ext + ";I;;" + "\n"
        #arq_csv.write(result_comp)
    print("-"*160)

    print("")
    print("")


    #VERIFICA SE TODOS OS ARQUIVOS EM ULTIMA_ENTREGA EXISTEM EM OBRIGACAO
    #arquivo2=ultima_entrega
    log("VERIFICA SE TODOS OS ARQUIVOS EM ULTIMA_ENTREGA EXISTEM EM " + configuracoes.OBRIGACAO )
    for arquivo2 in regs_ultima_entrega:
        
        
        
        
        print("")
        print("")
        log("verificando ",arquivo2)
        
        tipo = arquivo2.split(".")[0][-1]
        ext = arquivo2.split(".")[-1]

        #monta a mascara com o nome do arquivo ULTIMA_ENTREGA. Path e mascara completos.
        arquivo2mask = arquivo2.split(".")[0][:-4] + "???" + arquivo2.split(".")[0][-1] + "." + arquivo2.split(".")[1]
        #pega só o nome do arquivo ultima_entrega
        arquivo_ultima_entrega_temp = arquivo2mask.split("/")[-1]
        #pega só o path de ultima_entrega
        path_ultima_entrega_temp = arquivo2mask.replace(arquivo_ultima_entrega_temp,"")
        #monta só o path de obrigacao.
        path_obrigacao_temp = path_ultima_entrega_temp.replace("ULTIMA_ENTREGA",configuracoes.OBRIGACAO)
        #verifica quantos arquivos de ultima_entrega correspondem aos de obrigacao
        path = list(pathlib.Path(path_obrigacao_temp).glob(arquivo_ultima_entrega_temp))
        
        qo = len(path)
        msg = ""
        if ( qo > 1 ):
            msg = "Existem mais de 1 arquivo correspondente em OBRIGACAO para um em ULTIMA_ENTREGA. Existe mais de um sequencial para o mesmo volume em OBRIGACAO."
            result_comp = ";"+ str(arquivo2) + ";" + tipo + ";" + ext + ";" + "?" + ";" + msg + ";; \n"
            log("ERRO -  ", msg)
            arq_csv.write(result_comp)
            err = 111
        elif ( qo < 1 ):
            msg = "Não existe arquivo correspondente em OBRIGACAO ao arquivo em ULTIMA_ENTREGA."
            result_comp = ";"+ str(arquivo2) + ";" + tipo + ";" + ext + ";" + "?" + ";" + msg + ";; \n"
            log("ERRO -  ",msg)
            arq_csv.write(result_comp)
            err = 111
        elif (qo == 1):
            log("OK - Arquivo ULTIMA_ENTREGA:")
            log("    ====> ",arquivo2)
            log("Existe na pasta OBRIGACAO como:")
            log("    ====> ",path[0])




   #VERIFICA SE TODOS OS ARQUIVOS EM OBRIGACAO EXISTEM EM ULTIMA_ENTREGA
    log("VERIFICA SE TODOS OS ARQUIVOS EM " + configuracoes.OBRIGACAO + " EXISTEM EM ULTIMA_ENTREGA")
    for arquivo1 in registros:
        print("")
        print("")
        log("verificando ",arquivo1)
        
        tipo = arquivo1.split(".")[0][-1]
        ext = arquivo1.split(".")[-1]
        
        #monta a mascara com o nome do arquivo OBRIGACAO. Path e mascara completos.
        arquivo1mask = arquivo1.split(".")[0][:-4] + "???" + arquivo1.split(".")[0][-1] + "." + arquivo1.split(".")[1]
        #pega só o nome do arquivo obrigacao
        arquivo_obrigacao_temp = arquivo1mask.split("/")[-1]
        #pega só o path de obrigacao
        path_obrigacao_temp = arquivo1mask.replace(arquivo_obrigacao_temp,"")
        #monta só o path de ultima_entrega.
        path_ultima_temp = path_obrigacao_temp.replace(configuracoes.OBRIGACAO, "ULTIMA_ENTREGA")
        #verifica quantos arquivos de ultima_entrega correspondem aos de obrigacao
        path = list(pathlib.Path(path_ultima_temp).glob(arquivo_obrigacao_temp))
        qe = len(path)
        msg = ""
        if ( qe > 1 ):
            msg = "Existem mais de 1 arquivo correspondente em ULTIMA_ENTREGA para um em OBRIGACAO. Existe mais de um sequencial para o mesmo volume em ULTIMA_ENTREGA."
            result_comp = str(arquivo1) + ";;" + tipo + ";" + ext + ";?;" + msg + ";; \n"
            log("ERRO -  ", msg)
            arq_csv.write(result_comp)
            err = 111
        elif ( qe < 1 ):
            msg = "Não existe arquivo correspondente em ULTIMA_ENTREGA ao arquivo em OBRIGACAO."
            result_comp = str(arquivo1) + ";;" + tipo + ";" + ext + ";?;" + msg + ";; \n"
            log("ERRO -  ",msg)
            arq_csv.write(result_comp)
            err = 111
        elif (qe == 1):
            log("OK - Arquivo Obrigação:")
            log("    ====> ",arquivo1)
            log("Existe na pasta ULTIMA_ENTREGA como:")
            log("    ====> ",path[0])

            #SELECIONA A DUPLA DE ARQUIVOS PARA COMPARACAO
            a_comparar.append((arquivo1,str(path[0])))
         
            
    print("")
    print("")
    
    
    #COMPARA OS ARQUIVOS SELECIONADOS:
    for comparar in a_comparar:
        arquivo1 = comparar[0]
        arquivo2 = comparar[1]
        err = 0
        print("-"*160)
        log("Comparando os arquivos:")
        log("1 = '",arquivo1,"'")
        log("  com:")
        log("2 = '",arquivo2,"'")
        
       
        if (not os.path.isfile(arquivo1)):
            ret = 11
            continue

        if (not os.path.isfile(arquivo2)):
            ret = 22
            continue

        ##### COMPARA O CONTEUDO DOS ARQUIVOS...
        tipo = arquivo1.split(".")[0][-1]
        ext = arquivo1.split(".")[-1] 

        #transforma os arquivos arq1 e arq2 em listas...
        arq1 = open(arquivo1,"r", encoding=comum.encodingDoArquivo(arquivo1)).read().split("\n")
        arq2 = open(arquivo2,"r", encoding=comum.encodingDoArquivo(arquivo2)).read().split("\n")
        

        if (tipo  in ('C','D','I','M')):
            
            l1=len(arq1)
            l2=len(arq2)
            
            for x in range(0,l1):
                #RETIRAR 10  50 A  59 => CODIGO DO ITEM
                #RETIRAR  6 258 A 263 => ALIQUOTA PIS/PASEP
                #RETIRAR  6 275 A 280 => ALIQUOTA CONFINS
                #RETIRAR 32 300 A 331 => CODIGO AUTENTIFICACAO (331 = FINAL)
                if (tipo == "I"):
                    arq1[x]=arq1[x][0:49]+arq1[x][59:257]+arq1[x][263:274]+arq1[x][280:299]
            for y in range(0,l2):
                #RETIRAR 10  50 A  59 => CODIGO DO ITEM
                #RETIRAR  6 258 A 263 => ALIQUOTA PIS/PASEP
                #RETIRAR  6 275 A 280 => ALIQUOTA CONFINS
                #RETIRAR 32 300 A 331 => CODIGO AUTENTIFICACAO (331 = FINAL)
                if (tipo == "I"):
                    arq2[y]=arq2[y][0:49]+arq2[y][59:257]+arq2[y][263:274]+arq2[y][280:299]
            
            ##### COMPARA A QUANTIDADE DE REGISTROS
            if (l1 != l2):
                log("ERRO - ", arquivo1, " e ", arquivo2, " possuem quantidade de registros diferentes:")
                log("ERRO - ", arquivo1, " = ", len(arq1), " registros")
                log("ERRO - ", arquivo2, " = ", len(arq2), " registros")
                result_comp = arquivo1 + ";"+ arquivo2 + ";" + tipo + ";" + ext + ";" + "#" + ";" + "Quantidade de registros diferentes."+  arquivo1 + "=> " + str(l1) + " e " + arquivo2 + " =>" + str(l2) + ";;" + "\n"
                arq_csv.write(result_comp)
                ret = 55
                err = 1

            else:    
                ##### COMPARA O CONTEUDO DOS ARQUIVOS
                conjunto1 = set(arq1)
                diferencas = conjunto1.difference(arq2)
                ndif = len(diferencas)
                if ( ndif != 0 ):
                    log("ERRO - ", arquivo2, " e ", arquivo1, " possuem ", ndif, " diferença(s).")
                    print("")
                if (ndif > 0):
                    for x in range(0,l1):
                        if (arq1[x] != arq2[x]):
                            result_comp = arquivo1 + ";"+ arquivo2 + ";" + tipo + ";" + ext + ";" + "#" + ";" + "Conteudo diferente" + ";" + arq1[x] + ";" + arq2[x] + " \n"
                            arq_csv.write(result_comp)
                            ret = 55
                            err = 1
                else:
                    #log(arquivo1)
                    #log("e")
                    #log(arquivo2)
                    log("São iguais.")
                    print("")
                    result_comp = arquivo1 + ";"+ arquivo2 + ";" + tipo + ";" + ext + ";" + "=" + ";" + "Arquivos iguais" + ";"+ "''" + "; \n"
                    arq_csv.write(result_comp)

    return ret






def selectArquivos(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,PASTA,UFI="SP") :
    log("Gerando lista de arquivos existentes em "+ PASTA + " a comparar....")

    pasta = configuracoes.dir_base_comp
    
    #DEBUG# - para testar em produção:
    #pasta = pasta[68:]
    
    pasta = os.path.join(pasta,UFI,DT_MESANO_INICIO[8:10],DT_MESANO_INICIO[3:5],"TBRA",CC_FILIAL.replace("'",""),"SERIE")
    pasta_inicial = pasta
    
    lst_all_files = []
    lst_series = []

    if not os.path.isdir(pasta):
        #log("ERRO - PASTA:\n", pasta ,"\nnão existe.")
        #ret = 66
        os.mkdir(pasta)
    else:
        if (CC_SERIE =="''" ):
            lst_series = [f for f in listdir(pasta) if not isfile(join(pasta, f))]
        else:
            lst_series.append(CC_SERIE.replace("'",""))
      
        for serie_atu in lst_series:
            pasta = os.path.join(pasta_inicial,serie_atu,PASTA)
            if not os.path.isdir(pasta):
                #log("ERRO - PASTA:\n", pasta ,"\nnão existe.")
                #ret = 66
                os.mkdir(pasta)
            else:
                lst_files = [f for f in listdir(pasta) if isfile(join(pasta, f))]
                for file in lst_files:
                    lst_all_files.append(os.path.join(pasta,file))

    return(lst_all_files)

def inicializar() :
    ufi = "SP"
    ret = 0
    iei=filiaisi=diamesanoi=diamesanof=seriesi=False
 
#         addParametro(nomeParametro, identificador = None, descricao = '', obrigatorio = False, exemplo = None, default = False) : 
    comum.addParametro('CC_IE',None, 'Inscricao estadual a ser processada.', True, '"108383949112"')
    comum.addParametro('CC_FILIAL',None, 'Filial a serem processada.', True, '"9144"')
    comum.addParametro('DT_MESANO_INICIO',None, 'Mês e ano inicial, mês com dois di­gitos, ano com quatro di­gitos.', True, '"012015"')
    comum.addParametro('DT_MESANO_FIM',None, 'Mês e ano final, mês com dois di­gitos, ano com quatro di­gitos.', True, '"012015"')
    comum.addParametro('CC_SERIE',None, 'Série(s) a serem processadas. Se for "" serão consideradas todas.', True, '"U K , 1, C"')
   
    
    if not comum.validarParametros() :
        ret = 3
    else:
        iei        = comum.getParametro('CC_IE')            # Tem que ser válido != ""
        filiaisi   = comum.getParametro('CC_FILIAL')        # Pode ser "", 1 ou várias separadas por vírgula. 
        mesanoii   = comum.getParametro('DT_MESANO_INICIO') # Tem que ser válida no formato MMYYYY
        mesanofi   = comum.getParametro('DT_MESANO_FIM')    # Pode ser "" ou Tem que ser válida no formato MMYYYY
        seriesi    = comum.getParametro('CC_SERIE')         # Pode ser "", 1 ou várias separadas por vírgula.
    
        iei = iei.strip()
        
        if (iei == ""):
            ret = 1
            log("ERRO - IE não foi informada. IE INVALIDO!")
            
        for ca in iei:
            if not ca in ['0','1','2','3','4','5','6','7','8','9']:
                ret = 1
                log("ERRO - IE possui caracteres não numéricos. IE INVALIDO!")
                break
    
        if (mesanofi == ""):
            log("ATENÇÃO: - Não foi informado MMAAAA final, será considerado o mesmo inicial, ou seja: ",mesanoii )
            mesanofi = mesanoii
            
        diamesanoi = '01/'+mesanoii[0:2]+'/'+mesanoii[2:6]
        diamesanof = '01/'+mesanofi[0:2]+'/'+mesanofi[2:6]
              
        if (int(mesanoii[0:2]) < 1 or int(mesanoii[0:2]) > 12 or int(mesanofi[0:2]) < 1 or int(mesanofi[0:2]) > 12 ):
            ret = 99
            log("ERRO - Mes inicial informado é inválido!", " Foi informado ", mesanoii[0:2], " MES ANO INICIAL - INVALIDO!")
    
        if (int(mesanofi[0:2]) < 1 or int(mesanofi[0:2]) > 12 ):
            ret = 99
            log("ERRO - Mes final informado é inválido!", " Foi informado ", mesanofi[0:2] , ". MES ANO FINAL- INVALIDO!")
        
        filiaisv = ""
        for fi in (filiaisi.split(",")):
            filiaisv = filiaisv + "'" + fi.strip() + "',"
        filiaisi = filiaisv[0:len(filiaisv)-1]  
    
        seriesv = ""
        for se in (seriesi.split(",")):
            seriesv = seriesv + "'" + se.strip() + "',"
        seriesi = seriesv[0:len(seriesv)-1]  
        
    return (ret,iei,filiaisi,diamesanoi,diamesanof,seriesi,ufi)


if __name__ == "__main__":
    ret = 0
    ret,CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,UFI = inicializar()
    if (ret == 0 ): 
        #ret = processar(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,UFI)
        ret = selectPastas(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,UFI)
        print(ret,CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,UFI)
        if ( ret != 0) :
            log('ERRO no processamento ... Verifique. RC = ', ret)
    sys.exit(ret)



