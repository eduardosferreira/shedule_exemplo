#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: checklist.py
  CRIACAO ..: 30/11/2021
  AUTOR ....: Airton Borges da Silva Filho / KYROS Consultoria
  DESCRICAO :
      Parâmetros Input: Filial, Série (opcional), Ano Mes Inicio, Ano Mes Fim
      compara os arquivos nas pastas:
      


    ./comparativo_conv115.py filial serie ddmmyyyyinicial ddmmyyyyfinal
    ./comparativo_conv115.py "108383949112" "0001" "012015" "012015" ""
    
    
    
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
    
def processar(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,UFI) :
    ret = 0
    reg_obrigacao_tmp = selectArquivos(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,"OBRIGACAO_OLD",UFI)
    regs_ultima_entrega_tmp = selectArquivos(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,"ULTIMA_ENTREGA",UFI)
    amdhms = str(datetime.datetime.now().strftime('%Y%m%d-%H%M%S'))
    
    #filtra registros que interessam....
    registros=[]
    regs_ultima_entrega=[]
    
    for reg in reg_obrigacao_tmp:
        nom = reg.split(".")[-2]
        ext = reg.split(".")[-1]
        tipo = nom[-1]
        if ( (not ext.isdigit()) or (tipo not in ('C','D','I','M'))  ):
            continue
        else:
            registros.append(reg)
   
    for reg in regs_ultima_entrega_tmp:
        nom = reg.split(".")[-2]
        ext = reg.split(".")[-1]
        tipo = nom[-1]
        if ( (not ext.isdigit()) or (tipo not in ('C','D','I','M'))  ):
            continue
        else:
            regs_ultima_entrega.append(reg)
     
    qtde = len(registros)
    log("Quantidade de arquivos em 'OBRIGACAO_OLD' = ", qtde)
        
    if ( qtde < 1 ):
        log("ERRO - Não existem arquivos em 'OBRIGACAO_OLD' a serem comparados.")
        ret = 55
        
    #DEBUG#
    #print("ARQUIVOS EM 'OBRIGACAO_OLD'....:")
    #for a in registros:
    #    print("    - ", a)
    #print("-"*160)
    #    
    
    qtde = len(regs_ultima_entrega)
    log("Quantidade de arquivos em 'ULTIMA_ENTREGA' = ", qtde)

    if ( qtde < 1 ):
        log("ERRO - Não existem arquivos em 'ULTIMA_ENTREGA' a serem comparados.")
        ret = 55
        
    #DEBUG#
    #print("ARQUIVOS 'ULTIMA_ENTREGA'....:")
    #for a in regs_ultima_entrega:
    #    print("    - ", a)
    #print("-"*160)

    #VERIFICA SE TODOS OS ARQUIVOS EM OBRIGACAO_OLD EXISTEM EM ULTIMA_ENTREGA
    for arquivo1 in registros:
        arquivo2 = arquivo1.replace("OBRIGACAO_OLD", "ULTIMA_ENTREGA")


        if (not os.path.isfile(arquivo2)):

            #monta nome do arquivo CSV
            CC_SERIE = arquivo2.split("/")[-3]
            n_arq_s = "comparativo_" + UFI + "_" + CC_FILIAL.replace("'","") + "_" + CC_SERIE.replace("'","") + "_" + DT_MESANO_INICIO[6:10] + DT_MESANO_INICIO[3:5] + ".csv"
            lst_pasta_csv = arquivo2.split("/")[:-1]
            pasta_csv = ""
            for part_pasta_csv in lst_pasta_csv:
                pasta_csv = os.path.join(pasta_csv,part_pasta_csv)
            pasta_csv = os.path.join("/",pasta_csv)
            if not os.path.isdir(pasta_csv):
                os.mkdir(pasta_csv)
            arq_saida = os.path.join(pasta_csv,n_arq_s)
    
            log("ERRO - Existe o Arquivo:\n",arquivo1,"\n    mas não existe o arquivo:\n",arquivo2)
            ret = 181
            result_comp = "'"+"'"+ ";"+ arquivo1 + ";" + tipo + ";" + ext + ";" + "?" + ";" + "Arquivo equivalente ao REGERADO (NA PASTA OBRIGACAO_OLD) não existe na pasta ULTIMA_ENTREGA ou nomes são diferentes" + ";" + "''" + ";" + amdhms + "\n"
            gravar(arq_saida,result_comp)
            continue
        
    #VERIFICA SE TODOS OS ARQUIVOS EM ULTIMA_ENTREGA EXISTEM EM OBRIGACAO
    for arquivo2 in regs_ultima_entrega:
        arquivo1 = arquivo2.replace("ULTIMA_ENTREGA","OBRIGACAO_OLD")

        if (not os.path.isfile(arquivo1)):

            #monta nome do arquivo CSV
            CC_SERIE = arquivo2.split("/")[-3]
            n_arq_s = "comparativo_" + UFI + "_" + CC_FILIAL.replace("'","") + "_" + CC_SERIE.replace("'","") + "_" + DT_MESANO_INICIO[6:10] + DT_MESANO_INICIO[3:5] + ".csv"
            lst_pasta_csv = arquivo2.split("/")[:-1]
            pasta_csv = ""
            for part_pasta_csv in lst_pasta_csv:
                pasta_csv = os.path.join(pasta_csv,part_pasta_csv)
            pasta_csv = os.path.join("/",pasta_csv)
            if not os.path.isdir(pasta_csv):
                os.mkdir(pasta_csv)
            arq_saida = os.path.join(pasta_csv,n_arq_s)

            log("ERRO - Existe o Arquivo:\n",arquivo2,"\n    mas não existe o arquivo:\n",arquivo1)
            ret = 282
            result_comp = arquivo2 + ";"+ "'"+"'"+ ";" + tipo + ";" + ext + ";" + "?" + ";" + "Arquivo equivalente ao ULTIMA_ENTREGA não existe na pasta OBRIGACAO (REGERADO) ou nomes são diferentes" + ";" + "''" + ";" + amdhms + "\n"
            gravar(arq_saida,result_comp)
            continue
    
    for arquivo1 in registros:
        err = 0
        arquivo2 = arquivo1.replace("OBRIGACAO_OLD", "ULTIMA_ENTREGA")

        print("-"*160)
        log("Comparando os arquivos:")
        log("1 = '",arquivo1,"'")
        log("  com:")
        log("2 = '",arquivo2,"'")
         
        #monta nome do arquivo CSV
        CC_SERIE = arquivo2.split("/")[-3]
        n_arq_s = "comparativo_" + UFI + "_" + CC_FILIAL.replace("'","") + "_" + CC_SERIE.replace("'","") + "_" + DT_MESANO_INICIO[6:10] + DT_MESANO_INICIO[3:5] + ".csv"
        lst_pasta_csv = arquivo2.split("/")[:-1]
        pasta_csv = ""
        for part_pasta_csv in lst_pasta_csv:
            pasta_csv = os.path.join(pasta_csv,part_pasta_csv)
        pasta_csv = os.path.join("/",pasta_csv)
        if not os.path.isdir(pasta_csv):
            os.mkdir(pasta_csv)
        arq_saida = os.path.join(pasta_csv,n_arq_s)
        
        
        if (not os.path.isfile(arquivo1)):
#            log("ERRO - Arquivo1 '",arquivo1,"' não encontrado.")
#            result_comp = "'"+"'"+ ";"+ arquivo1 + ";" + tipo + ";" + ext + ";" + "?" + ";" + "Arquivo não encontrado" + ";" + "''" + ";" + amdhms + "\n"
#            arq_csv.write(result_comp)
            ret = 11
            continue

        if (not os.path.isfile(arquivo2)):
#            log("ERRO - Arquivo 2 '",arquivo2,"' não encontrado.")
#            result_comp = arquivo2 + ";"+ "'"+"'"+ ";" + tipo + ";" + ext + ";" + "?" + ";" + "Arquivo não encontrado" + ";" + "''" + ";" + amdhms + "\n"
#            arq_csv.write(result_comp)
            ret = 22
            continue

        ##### COMPARA O CONTEUDO DOS ARQUIVOS...

        #transforma os arquivos arq1 e arq2 em listas...
        arq1 = open(arquivo1,"r", encoding=comum.encodingDoArquivo(arquivo1)).read().split("\n")
        arq2 = open(arquivo2,"r", encoding=comum.encodingDoArquivo(arquivo2)).read().split("\n")
        
        conjunto1 = set(arq1)
        
        ##### COMPARA A QUANTIDADE DE REGISTROS
        if ( len(arq1) != len(arq2 )):
            log("ERRO - ", arquivo1, " e ", arquivo2, " possuem quantidade de registros diferentes:")
            log("ERRO - ", arquivo1, " = ", len(arq1), " registros")
            log("ERRO - ", arquivo2, " = ", len(arq2), " registros")
            result_comp = arquivo2 + ";"+ arquivo1 + ";" + tipo + ";" + ext + ";" + "#" + ";" + "Quantidade de registros diferentes" + ";" + "''" + ";" + amdhms + "\n"
            gravar(arq_saida,result_comp)

            ret = 55
            err = 1
            
        ##### COMPARA O CONTEUDO DOS ARQUIVOS
        diferencas = conjunto1.difference(arq2)
        ndif = len(diferencas)
        if ( ndif != 0 ):
            log("ERRO - ", arquivo1, " e ", arquivo2, " possuem ", ndif, " diferença(s).")

            for dif in diferencas:
                print(dif)
                result_comp = arquivo2 + ";"+ arquivo1 + ";" + tipo + ";" + ext + ";" + "#" + ";" + "Conteudo diferente" + ";" + "'" + dif + "'" + ";" + amdhms + "\n"
                gravar(arq_saida,result_comp)
            ret = 55
            err = 1
            
        if ( err == 0 ):
            result_comp = arquivo2 + ";"+ arquivo1 + ";" + tipo + ";" + ext + ";" + "=" + ";" + "Arquivos iguais" + ";"+ "''" + ";" + amdhms + "\n"
            gravar(arq_saida,result_comp)

    return ret

def selectArquivos(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,PASTA,UFI="SP") :
    log("Gerando lista de arquivos a comparar....")

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
 
    if not getattr(configuracoes, 'banco', False) :
        log("Erro falta variavel 'banco' no arquivo de configuração (.cfg).")
        ret = 1

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
        filiaisi   = comum.getParametro('CC_FILIAL')       # Pode ser "", 1 ou várias separadas por vírgula. 
        mesanoii   = comum.getParametro('DT_MESANO_INICIO') # Tem que ser válida no formato MMYYYY
        mesanofi   = comum.getParametro('DT_MESANO_FIM')    # Pode ser "" ou Tem que ser válida no formato MMYYYY
        seriesi    = comum.getParametro('CC_SERIE')        # Pode ser "", 1 ou várias separadas por vírgula.
    
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
        ret = processar(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,UFI)
        if ( ret != 0) :
            log('ERRO no processamento ... Verifique. RC = ', ret)
    sys.exit(ret)



