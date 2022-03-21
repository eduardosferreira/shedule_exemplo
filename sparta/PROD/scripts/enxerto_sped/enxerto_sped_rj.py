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
            Este script faz a geraçao da retificaçao do arquivo SPED FISCAL.
            
            VARIÁVEIS:  
            	<MMAAAA> =  MM: mes, numérico com dois dígitos. AAAA: ano, numérico com quatro dígitos.  ex: 012021
            	<IE>     =  Inscriçao estadual. numérico.  ex: 77452443
            	<UF> 	 =  Unidade Federativa, Estado. ex: RJ
            	<SSS>    =  Número sequencial com tres dígitos e com zeros a esquerda para completar os tres dígitos. ex: 001
            	<S/N>    =  A letra S ou a letra N.  ex: S
            
            O script enxerto_SPED faz o enxerto de blocos e registros a partir dos arquivos REGERADO e PROTOCOLADO, exceto do bloco 1400 que possui um script próprio.
            
            Arquivo Regerado = Arquivo SPED FISCAL gerado através da aplicaçao GF, após todas as correçoes efetuadas pelo Teshuvá
            
            Arquivo Protocolado = Arquivo SPED FISCAL entregue ao fisco pelo time tributário, pode conter ou nao a assinatura digital.
            
            As regras de escolha e alteraçoes dos blocos a serem inseridas no novo arquivo ENXERTADO estao detalhadas no documento: Discovery RJ - Enxerto - Requisitos Funcional - V11.doc 
            
            Para executar este script, será necessário fornecer os seguintes dados:
            
            <MMAAAA> =  MM: mes, numérico com dois dígitos. AAAA: ano, numérico com quatro dígitos.  ex: 012021
            <IE>     =  Inscriçao estadual. numérico.  ex: 77452443
            <S/N>    =  S=SIM ou N=NAO para forçar fazer o enxerto do bloco 1600 sempre partir do REGERADO,  independente de data.  
            			S:  Faz o enxerto a partir do REGERADO independente da data dos dados.  
            			N:  Faz o enxerto a partir do REGERADO se a data dos dados for menor ou igual a 072017 ou
            				Faz o enxerto a partir do PROTOCOLADO se a data dos dados for maior ou igual a 082017     
            
            Exemplo:	Enxerto_SPED_RJ.py 062020 77452443 S
            
            OBS:  	1 - Este script nao faz o enxerto do bloco 1400. Existe um outro script para o enxerto do bloco 1400.
            		2 - Os arquivos PROTOCOLADO e REGERADO devem existir e serem relativos aos mesmos dados, ou seja, mesma UF, IE, MM e AAAA.
            		3 - O arquivo PROTOCOLADO deve estar na pasta: /arquivos/SPED_FISCAL/PROTOCOLADOS/<UF>/<AAAA>/<MM>
            		4 - O nome do PROTOCOLADO deve ser: SPED_<MMAAAA>_<UF>_<IE>_PROT_V<SSS>.TXT
            		5 - O arquivo REGERADO deve estar na pasta: /arquivos/SPED_FISCAL/REGERADO/<UF>/<AAAA>/<MM>
            		6 - O nome do REGERADO deve ser: SPED_<MMAAAA>_<UF>_<IE>_REG_V<SSS>.TXT
            		7 - O arquivo ENXERTADO será criado na pasta: /arquivos/SPED_FISCAL/ENXERTADOS/<UF>/<AAAA>/<MM>
            		8 - O nome do ENXERTADO será: SPED_<MMAAAA>_<UF>_<IE>_ENX_V<SSS>.TXT, onde o <SSS> é o mesmo <SSS> do REGERADO.
----------------------------------------------------------------------------------------------
  HISTORICO : DOC v11 = 20210705 - Airton Borges da Silva Filho / Kyros Tec.
              20210816 - Adaptação para o novo Painel de execuções - Airton Borges da Silva Filho / Kyros Tec.
    - 22/02/2022 - Eduardo da Silva Ferreira - Kyros Tecnologia
            - [PTITES-1637] Padrão de diretórios do SPARTA
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
#import sql
comum.log.gerar_log_em_arquivo = False
#banco=sql.geraCnxBD(configuracoes)

#if (configuracoes.ambiente == 'DEV'):
#    dir_enxertados = dir_dev+dir_enxertados
#### PATRONIZACAO PARA O PAINEL DE EXECUCOES....






import datetime
import re
from pathlib import Path

global ret
global variaveis
global db
global arquivo_destino

toti = {}
totc = {}
g130p = []
g140p = []




toti['9900'] = 4

#DEBUG = True
DEBUG = False

filhos0200 = ('0205','0206','0210','0220')
filhos0150 = ('0175')

relatorio_erros = None
totais = {}


def contarLinhasArquivo(nome_arquivo):
    def blocks(files, size=65536):
        while True:
            b = files.read(size)
            if not b: break
            yield b
    encoding = encodingDoArquivo(nome_arquivo)
    with open(nome_arquivo, "r", encoding=encoding, errors='ignore') as f:
        return sum(bl.count("\n") for bl in blocks(f))

def formatNumero(numero): return '{:,}'.format(numero).replace(',', '.')


def encodingDoArquivo(path_arq) :
    global ret
    
    try :
        fd = open(path_arq, 'r', encoding='iso-8859-1')
        fd.read()
        fd.close()
    except :
        return 'utf-8'

    return 'iso-8859-1'


def retornaUFArquivo(path) :
    try :
        fd = open(path,'r') #, encoding=encodingDoArquivo(path))
        lin = fd.readline()
    except :
        fd = open(path,'r', encoding=encodingDoArquivo(path))
        lin = fd.readline()
    fd.close()
    if lin and lin.startswith('|0000|') :
        mes_ano = "%s_%s"%( lin.split('|')[4][2:4], lin.split('|')[4][4:] )

        return [lin.split('|')[9], mes_ano ] or [ False, False ]
    return False, False


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
        log("-"*100)
        log('ERRO:    Arquivo %s não está na pasta %s'%(mascara,diretorio))
        log("-"*100)
    return(nomearq)


def validauf(uf):
    return(True if (uf.upper() in ('AC','AL','AM','AP','BA','CE','DF','ES','GO','MA','MG','MS','MT','PA','PB','PE','PI','PR','RJ','RN','RO','RR','RS','SC','SE','SP','TO')) else False)

def ultimodia(ano,mes):
   return(31 if mes == 12 else datetime.date.fromordinal((datetime.date(ano,mes+1,1)).toordinal()-1).day)

def ies_existentes(mascara,diretorio):
    global ret
    
    qdade = 0
    ies = []
    directory = Path(diretorio)
    files = directory.glob(mascara)
    sorted_files = sorted(files, reverse=False)
    if sorted_files:
        log ("# Arquivos encontrados: ")
        for f in sorted_files:
            qdade = qdade + 1
            ie = str(f).split("_")[4]
            log("#   ",qdade, " => ", f, " IE = ", ie)
            try:
                ies.index(str(f).split("_")[4])
            except:
                ies.append(str(f).split("_")[4])
                continue
            
    else: 
        log('ERRO:    Arquivo %s não está na pasta %s'%(mascara,diretorio))
        ret=99
        return("")
    log("-"*100)
    return(ies)

def processar(ufi,mesanoi,mesi,anoi,iei,e1600):
    global ret

    anomesi = anoi+mesi
    nome_protocolado=""
    nome_regerado=""
    dir_base = SD + 'arquivos' + SD + 'SPED_FISCAL'
    dir_protocolados = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'SPED_FISCAL', 'PROTOCOLADOS', ufi, anoi, mesi) # [PTITES-1637] #
    dir_regerados = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'REGERADOS', ufi, anoi, mesi) # [PTITES-1637] #
 
 
 
 
 
 
    dir_dev = os.getcwd()
    
    dir_enxertados = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'ENXERTADOS',ufi,anoi,mesi) # [PTITES-1637] #
  
    # [PTITES-1637] # if (configuracoes.ambiente == 'DEV'):
    # [PTITES-1637] #    dir_enxertados = dir_dev+dir_enxertados
  





    mascara_regeradoi = "SPED_"+mesanoi+"_"+ufi+"_"+iei+"_REG*.txt"
    listadeies = ies_existentes(mascara_regeradoi,dir_regerados)
    
    for iee in listadeies:
        
        log("#")
        log("#")
        log("#")
        log("-"*100)
        log("#")
        log("#  INÍCIO do processamento para a IE ", iee)
        
        mascara_regerado = "SPED_"+mesanoi+"_"+ufi+"_"+iee+"_REG*.txt"
        mascara_protocolado = "SPED_"+mesanoi+"_"+ufi+"_"+iee+"_PROT*.txt"
        
        nome_regerado = nome_arquivo(mascara_regerado,dir_regerados)
        nome_protocolado = nome_arquivo(mascara_protocolado,dir_protocolados)
         
        if ((nome_regerado == "") or (nome_protocolado == "")):
            log("-"*100)

            log("ERRO: Não foi processado o ENXERTO para o arquivo regerado    = ", nome_regerado)
            log("ERRO: Não foi processado o ENXERTO para o arquivo protocolado = ", nome_protocolado)
            log("-"*100)
            ret=99
        else:
                
            ### prepara saida ENXERTADO
            if (str(nome_regerado).count("_") == 6):
                versao_enxertado = "_"+(str(nome_regerado).split(".")[0]).split("_")[6]
            else:
                versao_enxertado = ""







            # [PTITES-1637] # if (configuracoes.ambiente == 'DEV'):
            # [PTITES-1637] #     dir_enxertados = dir_dev+dir_enxertados                
                
  




            nome_enxertado  = os.path.join(dir_enxertados, "SPED_"+mesanoi+"_"+ufi+"_"+iee+"_ENX"+versao_enxertado+ ".txt")
        
            if not os.path.isdir(dir_enxertados) :
                os.makedirs(dir_enxertados)

            log("#")
            log("-"*100)
            log("#  Arquivos a serem processados:")
            log("#     Arquivo protocolado = ", nome_protocolado)
            log("#     Arquivo regerado    = ", nome_regerado)
            log("#     Arquivo enxertado   = ", nome_enxertado)
            log("#     1600 do Reg ou Prot = ", e1600)
            log("-"*100)


            log("#  Preparando os registros 1600...")
            if (e1600 == "R"):
                arq1600 = nome_regerado
            else:
                arq1600 = nome_protocolado
            lista1600 = []
            lista1600 = prepara1600(arq1600)
            log("#  Fim do Preparo dos registros 1600.")

            if (len(lista1600) == 0):
                log('-'* 160)
                log('####')
                if (e1600 == "R"):
                    log('ERRO - O ARQUIVO REGERADO NÃO POSSUI REGISTROS 1600, O ENXERTO DO 1600 NÃO FOI REALIZADO.')
                    ret = 88
                else: 
                    log('ERRO - O ARQUIVO PROTOCOLADO NÃO POSSUI REGISTROS 1600, O ENXERTO DO 1600 NÃO FOI REALIZADO.')
                    ret = 88
                log('####')
                log('-'* 160)
               
            log("#  Contando os registros 1400...")
            q1400 = conta1400(nome_regerado)
            log("#  Quantidade de 1400 a enxertar = ", q1400)       
            
            if (processaDiretorio(nome_protocolado, nome_regerado, nome_enxertado, dir_protocolados, dir_regerados, dir_enxertados, lista1600, q1400, anomesi) == False):
                ret = 99
 
        log("-"*100)
        log("#")
        log("# FIM do processamento para a IE ", iee)
        log("#")
        log("-"*100)
        log("#")
        log("#")
        log("#")

    return(ret)

def prepara1600(arq1600):
    bloco1600 = []
    enc = encodingDoArquivo(arq1600)
    temp = open(arq1600, 'r', encoding=enc, errors='ignore')
    for lin in temp:
        if lin.startswith('|1600|') :
            bloco1600.append(lin)
    temp.close() 
    return(bloco1600)

def conta1400(arq1400):
    qdade1400 = 0
    enc = encodingDoArquivo(arq1400)
    temp = open(arq1400, 'r', encoding=enc, errors='ignore')
    for lin in temp:
        if lin.startswith('|1400|') :
            qdade1400 = qdade1400 + 1
    temp.close() 
    return(qdade1400)

    
def parametros():
    global ret
    ufi = ""
    mesanoi = ""
    iei = "*"
    mesi = ""
    anoi = "" 
    ret = 0
    ufi = "RJ"
     
    if (len(sys.argv) == 4
        and len(sys.argv[1])==6  
        and int(sys.argv[1][:2])>0 
        and int(sys.argv[1][:2])<13
        and int(sys.argv[1][2:])<=datetime.datetime.now().year
        and int(sys.argv[1][2:])>(datetime.datetime.now().year)-50
        ):
     
        iei=sys.argv[2].upper()
        iei = re.sub('[^0-9]','',iei)
        if ( (iei == "") or (iei == "''") or (iei == '""') or (int("0"+iei) == 0)):
            iei = "*"
             
    else :
        log('ERRO - Erro nos parametros do script.')
        log('####')
        log('#### Exemplo de como deve ser :')
        log('####      %s <MMAAAA> <IE> <S|N>'%(sys.argv[0]))
        log('####')
        log('#### Onde')
        log('####      <MMAAAA> = mês e ano. Ex: Para junho de 2020 informe 062020')
        log('####      <IE> =     Inscição Estadual.')
        log('####      <S|N> =    Se informado "S", realiza enxerto do 1600 do REGERADO independente de data.')
        log('####                 Se informado "N", realiza enxerto do 1600 do REGERADO até 07/2017 ou do PROTOCOLADO para 08/2017 ou após.')
        log('####')
        log('#### Portanto, se o mes = 06 e o ano = 2020, IE = 77452443 e deseja realizar o enxerto do 1600 independente de data, o comando correto deve ser :')  
        log('####      %s 062020 77452443 S'%(sys.argv[0]))  
        log('####')
        log("-" * 100)
        log("")
        log("Retorno = 99") 
        ret = 99
        return(False,False,False,False,False,False)

    mesanoi = sys.argv[1].upper()
    mesi    = sys.argv[1][:2].upper()
    anoi    = sys.argv[1][2:].upper()
    e1600i  = sys.argv[3].upper()
    
    anomesi = int(anoi+mesi)
    
    if (anomesi < 201708):
        e1600 = "R"
    else:
        e1600 = "P"
        
    if (e1600i == "S"):
        e1600 = "R"
    
 
    return(ufi,mesanoi,mesi,anoi,iei,e1600)


def processaDiretorio(nome_protocolado, nome_regerado, nome_enxertado, path_protocolados, path_regerados, path_enxertados, l1600, q1400, anomesi) :
    global relatorio_erros 
    
    q1600 = len(l1600)
    
    path_nomeArquivoProtocolado = nome_protocolado
    path_nomeArquivoRegerado = nome_regerado
    path_nomeArquivoEnxertado  = nome_enxertado
  
    #### MONTA OS BLOCOS G140 E G130
    #### MONTA OS BLOCOS G140 E G130
    #### MONTA OS BLOCOS G140 E G130
    #### MONTA OS BLOCOS G140 E G130
    #### MONTA OS BLOCOS G140 E G130


    
    pg130 = []
    pg140 = []
    pg1402 = []
    
    b0150p = []
    b0190p = []
    b0200p = []
    
    bp = []
    
    log("Parte 01 / 17 - Montando lista condicional com blocos G130, G140 e 1600 para atender condição dos blocos 0150, 0190 e 0200...")
   
    encP = encodingDoArquivo(path_nomeArquivoProtocolado)
    tempP = open(path_nomeArquivoProtocolado, 'r', encoding=encP, errors='ignore')
    
    for lP in tempP:

        if lP.startswith('|G130|') :
            if ( len(lP.split('|')) > 2 ):
                cod = lP.split('|')[2]
                if (not cod in pg130):
                    pg130.append(cod)

        if lP.startswith('|1600|') :
            if ( len(lP.split('|')) > 2 ):
                cod = lP.split('|')[2]
                if (not cod in pg130):
                    pg130.append(cod)
                
        if lP.startswith('|G140|') :
            if ( len(lP.split('|')) > 3 ):
                cod  = lP.split('|')[3]
                if (not cod in pg1402):
                    pg1402.append(cod)

        if lP.startswith('|G140|') :
            if ( len(lP.split('|')) > 5 ):
                cod  = lP.split('|')[5]
                if (not cod in pg140):
                    pg140.append(cod)

                
    tempP.close() 
    
    if DEBUG: 
        log("########## DEBUG ##########")
        log("########## DEBUG ##########")
        log("pg130  = ", pg130)
        log("pg140  = ", pg140)
        log("pg1402 = ", pg1402)
        log("########## DEBUG ##########")
        log("########## DEBUG ##########")    
    
    
    
    
    
    log("Parte 02 / 17 - Selecionando os registros dos blocos 0150, 0190 e 0200 que atendem a condição de existirem nos blocos G130 e G140. ..")
    
    tempP = open(path_nomeArquivoProtocolado, 'r', encoding=encP, errors='ignore')

    for lP in tempP:
        if lP.startswith('|0150|') :
            if ( len(lP.split('|')) > 2 ):
                cod = lP.split('|')[2]
                if (cod in pg130):
                    b0150p.append(cod)
                    bp.append('|' + lP.split('|')[1] + '|' + lP.split('|')[2] + '|')
        elif lP.startswith('|0190|') :
            if ( len(lP.split('|')) > 2 ):
                cod = lP.split('|')[2]
                if (cod in pg140):
                    b0190p.append(cod)
                    bp.append('|' + lP.split('|')[1] + '|' + lP.split('|')[2] + '|')
        elif lP.startswith('|0200|') :
            if ( len(lP.split('|')) > 2 ):
                cod = lP.split('|')[2]
                if (cod in pg1402):
                    b0200p.append(cod)
                    bp.append('|' + lP.split('|')[1] + '|' + lP.split('|')[2] + '|')
                
    tempP.close() 
    

    log("Parte 03 / 17 - Gerando o arquivo com o enxerto de todos os blocos...")



    #### VERIFICAÇOES
    #### VERIFICAÇOES
    #### VERIFICAÇOES
    #### VERIFICAÇOES

    #Abre arquivo PROTOCOLADO para processamento principal    
    encP = encodingDoArquivo(path_nomeArquivoProtocolado)
    arquivoP = open(path_nomeArquivoProtocolado, 'r', encoding=encP, errors='ignore')
    num_linhaP = 0

    #ler até encontrar o primeiro registro....
    for linhaP in arquivoP:
        num_linhaP = num_linhaP + 1
        if linhaP.startswith('|0000|') :
            break



    #Reabre arquivo REGERADO para processamento principal 
    encR = encodingDoArquivo(path_nomeArquivoRegerado)
    arquivoR = open(path_nomeArquivoRegerado, 'r', encoding=encR, errors='ignore')
    num_linhaR = 0

    #ler até encontrar o primeiro registro....
    for linhaR in arquivoR:
        num_linhaR = num_linhaR + 1
        if linhaR.startswith('|0000|') :
            break


    #Verifica se os arquivos são do mesmo periodo.
    if (linhaP.split('|')[4] != linhaR.split('|')[4]) or (linhaP.split('|')[5] != linhaR.split('|')[5]) :
        log("#"*80)
        log('### ERRO - Arquivos com periodos de dados diferentes ... Verifique !!')
        log("#"*80)
        return False


    #### INICIO DA GERACAO DO ARQUIVO ENXERTADO
    #### INICIO DA GERACAO DO ARQUIVO ENXERTADO
    #### INICIO DA GERACAO DO ARQUIVO ENXERTADO
    #### INICIO DA GERACAO DO ARQUIVO ENXERTADO
    #### INICIO DA GERACAO DO ARQUIVO ENXERTADO


            
    #Cria o arquivo enxertado
    arquivoE = open(path_nomeArquivoEnxertado, 'w', encoding=encR, errors='ignore')
    numlinE = 0
    
    

    ##### BLOCO 4
    ##### BLOCO 4
    ##### BLOCO 4
    ##### BLOCO 4
    ##### BLOCO 4
    ##### BLOCO 4
    ###############################################################################   
    log("Parte 04 / 17 - Blocos >= 0000 a < 0150...")
    #gravar linha |0000| no Enxertado
    numlinE = gravar(arquivoE, linhaR[:], numlinE)




    #gravar regerado no ArquivoE até encontrar 0150 ou maior
    #De acordo 06/05/2021 - gravar 
    for linhaR in arquivoR:
        num_linhaR = num_linhaR + 1
        if (linhaR.split('|')[1] < '0150') :
            numlinE = gravar(arquivoE, linhaR[:], numlinE)
        else:
            break
        #endif
    #endfor

    #posiciona protocolado na primeira linha >=150
    for linhaP in arquivoP:
        num_linhaP = num_linhaP + 1
        if (linhaP.split('|')[1] < '0150') :
            continue
        else:
            break
        #endif
    #endfor
    ###############################################################################


    ###### BLOCO 5
    ###### BLOCO 5
    ###### BLOCO 5
    ###### BLOCO 5
    ###### BLOCO 5
    ###### BLOCO 5
    ###### BLOCO 5
    ###############################################################################
    log("Parte 05 / 17 - Blocos >= 0150 a <= 0200...")
    #gravar linha >=150 e <=200 no enxertado, pega do regerado. Se a do protocolado não existir no regerado, pega se existir no G130 ou G140 do protocolado.
    valorP = linhaP.split('|')[1]
    valorR = linhaR.split('|')[1]

    while (valorR >= '0150' and valorR <= '0220' and valorP >= '0150' and valorP <= '0220'):
        #        c0150 = linhaP.split('|')[2]
    #        c0190 = linhaP.split('|')[2]
    #        c0200 = linhaP.split('|')[2]
    #        bloco = linhaP.split('|')[1]
        cadgP = '|' + linhaP.split('|')[1] + '|' + linhaP.split('|')[2] + '|'
        cadgR = '|' + linhaR.split('|')[1] + '|' + linhaR.split('|')[2] + '|'

        if (cadgP < cadgR) :
            if ( (cadgP in bp) ):
                numlinE = gravar(arquivoE, linhaP[:], numlinE)
                for linhaP in arquivoP:
                    valorP = linhaP.split('|')[1]
                    break
                #####INICIO  GRAVA FILHOS DE PROTOCOLADO
                while ( (valorP in filhos0150) or (valorP in filhos0200) ):
                    numlinE = gravar(arquivoE, linhaP[:], numlinE)
                    for linhaP in arquivoP:
                        valorP = linhaP.split('|')[1]
                        break
                ##### FIM GRAVA FILHOS DE PROTOCOLADO
            else:
                #### PULA PARA O PROXIMO PROTOCOLADO
                for linhaP in arquivoP:
                    valorP = linhaP.split('|')[1]
                    break
                #####INICIO PULA FILHOS DE PROTOCOLADO
                while ( (valorP in filhos0150) or (valorP in filhos0200) ):
                    for linhaP in arquivoP:
                        valorP = linhaP.split('|')[1]
                        break
                ##### FIM PULA FILHOS DE PROTOCOLADO                    
        
        if (cadgP > cadgR) :
            numlinE = gravar(arquivoE, linhaR[:], numlinE)
            for linhaR in arquivoR:
                valorR = linhaR.split('|')[1]
                break
            #####INICIO  GRAVA FILHOS DE REGERADO
            while ( (valorR in filhos0150) or (valorR in filhos0200) ):
                numlinE = gravar(arquivoE, linhaR[:], numlinE)
                for linhaR in arquivoR:
                    valorR = linhaR.split('|')[1]
                    break
            ##### FIM GRAVA FILHOS DE REGERADO                    
        
        if (cadgP == cadgR) :
            if ( (cadgP in bp) ):
                ##### GRAVA O PROTOCOLADO
                numlinE = gravar(arquivoE, linhaP[:], numlinE)
                for linhaP in arquivoP:
                    valorP = linhaP.split('|')[1]
                    break
                #####INICIO  GRAVA FILHOS DE PROTOCOLADO
                while ( (valorP in filhos0150) or (valorP in filhos0200) ):
                    numlinE = gravar(arquivoE, linhaP[:], numlinE)
                    for linhaP in arquivoP:
                        valorP = linhaP.split('|')[1]
                        break
                ##### FIM GRAVA FILHOS DE PROTOCOLADO
                
                ##### INICIO POSICIONA O REGERADO NO PROXIMO
                for linhaR in arquivoR:
                    valorR = linhaR.split('|')[1]
                    break
                while ( (valorR in filhos0150) or (valorR in filhos0200) ):
                    for linhaR in arquivoR:
                        valorR = linhaR.split('|')[1]
                        break
                ##### FIM POSICIONA O REGERADO NO PROXIMO
            else: 
                ##### GRAVA O REGERADO
                numlinE = gravar(arquivoE, linhaR[:], numlinE)
                for linhaR in arquivoR:
                    valorR = linhaR.split('|')[1]
                    break
                #####INICIO  GRAVA FILHOS DE REGERADO
                while ( (valorR in filhos0150) or (valorR in filhos0200) ):
                    numlinE = gravar(arquivoE, linhaR[:], numlinE)
                    for linhaR in arquivoR:
                        valorR = linhaR.split('|')[1]
                        break
                ##### FIM GRAVA FILHOS DE REGERADO        
            
                #### PULA PARA O PROXIMO PROTOCOLADO
                for linhaP in arquivoP:
                    valorP = linhaP.split('|')[1]
                    break
                #####INICIO PULA FILHOS DE PROTOCOLADO
                while ( (valorP in filhos0150) or (valorP in filhos0200) ):
                    for linhaP in arquivoP:
                        valorP = linhaP.split('|')[1]
                        break
                ##### FIM PULA FILHOS DE PROTOCOLADO                              
    #endwhile
    
    while (valorR >= '0150' and valorR <= '0220'):
        numlinE = gravar(arquivoE, linhaR[:], numlinE)
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break
        #####INICIO  GRAVA FILHOS DE REGERADO
        while ( (valorR in filhos0150) or (valorR in filhos0200) ):
            numlinE = gravar(arquivoE, linhaR[:], numlinE)
            for linhaR in arquivoR:
                valorR = linhaR.split('|')[1]
                break
        ##### FIM GRAVA FILHOS DE REGERADO         


    while (valorP >= '0150' and valorP <= '0220'):
        if ( (cadgP in bp) ):
            ##### GRAVA O PROTOCOLADO
            numlinE = gravar(arquivoE, linhaP[:], numlinE)
            for linhaP in arquivoP:
                valorP = linhaP.split('|')[1]
                break
            #####INICIO  GRAVA FILHOS DE PROTOCOLADO
            while ( (valorP in filhos0150) or (valorP in filhos0200) ):
                numlinE = gravar(arquivoE, linhaP[:], numlinE)
                for linhaP in arquivoP:
                    valorP = linhaP.split('|')[1]
                    break
            ##### FIM GRAVA FILHOS DE PROTOCOLADO
        else: 
            for linhaP in arquivoP:
                valorP = linhaP.split('|')[1]
                break
            #####INICIO PULA FILHOS DE PROTOCOLADO
            while ( (valorP in filhos0150) or (valorP in filhos0200) ):
                for linhaP in arquivoP:
                    valorP = linhaP.split('|')[1]
                    break
            ##### FIM PULA FILHOS DE PROTOCOLADO           
    ###############################################################################
    # =============================================================================


    ###### BLOCO 6
    ###### BLOCO 6
    ###### BLOCO 6
    ###### BLOCO 6
    ###### BLOCO 6
    ###### BLOCO 6
    ###############################################################################
    log("Parte 06 / 17 - Blocos do protocolado <= 0200 a < 0400...")

    #Posiciona o PROTOCOLADO no 0300
    while valorP < '0300':
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break       
    # Grava do protocolado até chegar no 0400
    while valorP < '0400':
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
    ###############################################################################





    ###### BLOCO 7
    ###### BLOCO 7
    ###### BLOCO 7
    ###### BLOCO 7
    ###### BLOCO 7
    ###### BLOCO 7
    ###############################################################################
    log("Parte 07 / 17 - Blocos do REGERADO >= 0400 a <= 0460...")
    #Posiciona o REGERADO no 0400
    while valorR < '0400':
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break
    #Grava do REGERADO até acabar o 0460
    while valorR <= '0460':
        numlinE = gravar(arquivoE, linhaR[:], numlinE)
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break
    ###############################################################################





    ###### BLOCO 8
    ###### BLOCO 8
    ###### BLOCO 8
    ###### BLOCO 8
    ###### BLOCO 8
    ###### BLOCO 8
    ###############################################################################
    log("Parte 08 / 17 - Blocos do PROTOCOLADO 0500 E DO REGERADO os 0500 que não existam no PROTOCOLADO...")
    #Posiciona o REGERADO no 0500
    while valorR < '0500':
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break
    #Posiciona o PROTOCOLADO no 0500
    while valorP < '0500':
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break       
    #Grava o 0500 do PROTOCOLADO e inclui os do REGERADO que não existam no PROTOCOLADO
    valorP = linhaP.split('|')[1]
    valorR = linhaR.split('|')[1]

    while (valorP == '0500' or valorR == '0500'):
        cadgP = '|' + linhaP.split('|')[1] + '|' + linhaP.split('|')[2] + '|'
        cadgR = '|' + linhaR.split('|')[1] + '|' + linhaR.split('|')[2] + '|'
        
        if (cadgP < cadgR) :
            numlinE = gravar(arquivoE, linhaP[:], numlinE)
            for linhaP in arquivoP:
                valorP = linhaP.split('|')[1]
                break

        if (cadgP > cadgR) :
            numlinE = gravar(arquivoE, linhaR[:], numlinE)
            for linhaR in arquivoR:
                valorR = linhaR.split('|')[1]
                break

        if (cadgP == cadgR) :
            numlinE = gravar(arquivoE, linhaP[:], numlinE)
            for linhaP in arquivoP:
                valorP = linhaP.split('|')[1]
                break
            for linhaR in arquivoR:
                valorR = linhaR.split('|')[1]
                break
    #endwhile
    ###############################################################################





    ###### BLOCO 9
    ###### BLOCO 9
    ###### BLOCO 9
    ###### BLOCO 9
    ###### BLOCO 9
    ###### BLOCO 9
    ###############################################################################
    #Posiciona o PROTOCOLADO no 0600
    log("Parte 09 / 17 - Blocos do PROTOCOLADO >= 0600 a <= 0990...")
    while valorP < '0600':
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break       
    #Grava do PROTOCOLADO os de 0600 até acabar o 0990
    while ( valorP <= '0990' ):  
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
    ###############################################################################






    ###### BLOCO 10
    ###### BLOCO 10
    ###### BLOCO 10
    ###### BLOCO 10
    ###### BLOCO 10
    ###### BLOCO 10
    ###############################################################################
    log("Parte 10 / 17 - Blocos do REGERADO > 0990 a <= D990...")
    #Posiciona o REGERADO NO próximo após o 0990
    while valorR <= '0990':
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break       
    #grava do regerado até o fim do D990
    while valorR <= 'D990':
        if valorR.isdigit():
            break
        numlinE = gravar(arquivoE, linhaR[:], numlinE)
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break
    ###############################################################################







    ###### BLOCO 11
    ###### BLOCO 11
    ###### BLOCO 11
    ###### BLOCO 11
    ###### BLOCO 11
    ###### BLOCO 11
    ###############################################################################
    log("Parte 11 / 17 - Blocos do PROTOCOLADO > D990 e <= 1001...")
    #posiciona o protocolado no > D990
    while valorP <= 'D990':
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
    #grava do protocolado até o 1001
    #    grava os iniciados com letra....
    while valorP[0].isdigit() == False:
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
    ###############################################################################




    ###### BLOCO 12
    ###### BLOCO 12
    ###### BLOCO 12
    ###### BLOCO 12
    ###### BLOCO 12
    ###### BLOCO 12
    ###############################################################################
    log("Parte 12 / 17 - Blocos do PROTOCOLADO 1001 e análise de acordo com o 1600...")
    
    R1001 = ""
    P1001 = "" 
    
    #posiciona o REGERADO no 1001
    while valorR[0].isdigit() == False:
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break

    #obtem do protocolado o 1001
    while ( valorP == '1001' ):  
        P1001 = linhaP[:]
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break

    #Obtem do regerado o 1001
    while ( valorR == '1001' ):  
        R1001 = linhaR[:]
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break            

    if (R1001 == P1001):
        if (len(R1001) > 0 ):
            numlinE = gravar(arquivoE, R1001, numlinE)         
    else:
       R1001 = '|1001|0|' 
       numlinE = gravar(arquivoE, R1001, numlinE)
       
    if (DEBUG):
        log("Registro 1001 = ", R1001)
       

    ###############################################################################






    ###### BLOCO 13
    ###### BLOCO 13
    ###### BLOCO 13
    ###### BLOCO 13
    ###### BLOCO 13
    ###### BLOCO 13
    ###############################################################################
    log("Parte 13 / 17 - Blocos do PROTOCOLADO = 1010 e análise de acordo com o 1600 e 1400...")
    #posiciona o protocolado no 1010
    while valorP < '1010':
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
    #posiciona o regerado no 1010
    while valorR < '1010':
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break
    #Monta o novo registro 1010
    if(q1600 > 0):
        t1600 = "S"
    else: 
        t1600 = "N"
        
    if(q1400 > 0):
        t1400 = "S"
    else: 
        t1400 = "N"
        
    nlR = "|1010|N|N|N|N|"+t1400+"|N|"+t1600+"|N|N|N|N|N|" + "\n"    
      
    if (anomesi < '201901'):
        nlR = "|1010|N|N|N|N|"+t1400+"|N|"+t1600+"|N|N|" + "\n"    
    if (anomesi > '201912'):
        nlR = "|1010|N|N|N|N|"+t1400+"|N|"+t1600+"|N|N|N|N|N|N|" + "\n"
    
    numlinE = gravar(arquivoE, nlR, numlinE)
    ###############################################################################


    ###### BLOCO 14
    ###### BLOCO 14
    ###### BLOCO 14
    ###### BLOCO 14
    ###### BLOCO 14
    ###### BLOCO 14
    ###############################################################################
    log("Parte 14 / 17 - Bloco 1400 do REGERADO...")
    #enxerta 1400

    while valorR[0].isdigit() == False:
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break

    while valorR < '1400':
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break
    #grava do REGERADO o 1400
    while ( valorR == '1400' ):  
        numlinE = gravar(arquivoE, linhaR[:], numlinE)
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break
    
    ###############################################################################









    ###### BLOCO 15
    ###### BLOCO 15
    ###### BLOCO 15
    ###### BLOCO 15
    ###### BLOCO 15
    ###### BLOCO 15
    ###############################################################################
    log("Parte 15 / 17 - Bloco 1600...")
    #enxerta 1600
    if (q1600 > 0 ):
        for linha1600 in l1600:
            numlinE = gravar(arquivoE, linha1600[:], numlinE) 
    ###############################################################################





    ###### BLOCO 16
    ###### BLOCO 16
    ###### BLOCO 16
    ###### BLOCO 16
    ###### BLOCO 16
    ###### BLOCO 16
    ###############################################################################
    log("Parte 16 / 17 - Blocos do PROTOCOLADO > 1600 ATÉ 9900...")
    #posiciona o PROTOCOLADO no > 1600
    while valorP <= '1600':
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break

    #grava do PROTOCOLADO até 0 9900
    while valorP != '9900':
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
    ###############################################################################





        
        
    #CONTAGEM DOS REGISTROS    
    #CONTAGEM DOS REGISTROS    
    #CONTAGEM DOS REGISTROS    
    #CONTAGEM DOS REGISTROS    
    #CONTAGEM DOS REGISTROS    
    ###############################################################################
    log("Parte 17 / 17 - Blocos TOTALIZADORES até o fim 9999...")

    totc['9990'] = 1
    totc['9999'] = 1
    totc['9900'] = 1

    for soma in totc:
        linhaf = "|9900|" + soma + "|" + str(totc[soma]) + "|\n"
        numlinE = gravar(arquivoE, linhaf, numlinE)
  
    toti['9990'] = toti['9990'] + 1
    
    linhaf = '|9990|' + str(toti['9990']) + '|\n'
    numlinE = gravar(arquivoE, linhaf, numlinE)

    linhaf = '|9999|' + str(numlinE + 1) + '|\n'
    numlinE = gravar(arquivoE, linhaf, numlinE)
    #    log("linhaf = ", linhaf)


    arquivoP.close()
    arquivoR.close()
    arquivoE.close()

    return True



def gravar(arquivo, dado, contador):
    contador = contador + 1
    chave = dado.split('|')[1]
    vi = chave[0]
    chaveti = vi + '990'

    #total por inicial
    if (chaveti in toti):
        toti[chaveti] = toti[chaveti] + 1
    else:
    #        log("Nova chave toti = ", chaveti)
        toti[chaveti] = 1
        
    #total por chave
    if (chave in totc):
        totc[chave] = totc[chave] + 1
    else:
    #        log("Nova chave totc = ", chave)
        totc[chave] = 1

    if ((chave[1:] == '990' )):
        dado = '|'+ str(chaveti) + '|'+ str(toti[chaveti]) + '|\n'

    #grava no destino
    arquivo.write(dado)
    return (contador)

if __name__ == "__main__":    
    global ret
    
    log('#'*100)
    log("# ")  
    log("#  - INICIO - ENXERTO_SPED_RJ VERSÃO 20210615 DOC V12")
    log("# ")
    log('#'*100)
    ret = 0
    retorno = parametros()

    if DEBUG:
        log("retorno de parametros() = ", retorno)
    
    ufi = retorno[0]
    if (retorno[0] != False):
        ret     = 0
        ufi     = retorno[0]
        mesanoi = retorno[1]
        mesi    = retorno[2]
        anoi    = retorno[3]
        iei     = retorno[4]
        e1600   = retorno[5]
        log("-"*100)
        log("# Processando ENXERTO SPED para os seguintes parâmetros:")
        log("#    UF    = ",retorno[0])
        log("#    MÊS   = ",retorno[2])
        log("#    ANO   = ",retorno[3])
        log("#    IE    = ",retorno[4])
        log("#    1600  = ",retorno[5])
        log("-"*100)
        
        if DEBUG:
            input("Continua e realiza o enxerto ?")
        
        ret = processar(ufi,mesanoi,mesi,anoi,iei,e1600)
    else:
        ret = 99

#    if(ret != 0):
#        log("ERRO - VERIFIQUE AS MENSAGENS ANTERIORES PARA IDENTIFICAR O ERRO. ",ret)
        
    log('#'*100)
    log("# ")  
    log("#  - FIM - ENXERTO_SPED_RJ VERSÃO 20210615 DOC V12")
    log("# ")
    log("#"*100)

    log("Codigo de saida = ",ret)

    sys.exit(ret)
  
 
