#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: Enxerto_SPED_PE
  CRIACAO ..: 12/01/2020
  AUTOR ....: Airton Borges da Silva Filho / KYROS Consultoria
  DESCRICAO : 
----------------------------------------------------------------------------------------------
  HISTORICO : 
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
import atexit
import re
#import glob
from pathlib import Path
global ret


global variaveis
#global db
global arquivo_destino



toti = {}
totv = {}


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


# =============================================================================
def encodingDoArquivo(path_arq) :
    global ret
    
    try :
        fd = open(path_arq, 'r', encoding='iso-8859-1')
        fd.read()
        fd.close()
    except :
        return 'utf-8'

    return 'iso-8859-1'

# =============================================================================

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
        log('#### ERRO:    Arquivo %s não está na pasta %s'%(mascara,diretorio))
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
        log('#### ERRO:    Arquivo %s não está na pasta %s'%(mascara,diretorio))
        ret=99
        return("")
    log("-"*100)
    return(ies)

def processar(ufi,mesanoi,mesi,anoi,iei):
    global ret
    
    nome_protocolado=""
    nome_regerado=""
    dir_base = SD + 'arquivos' + SD + 'SPED_FISCAL'
    dir_protocolados = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'SPED_FISCAL', 'PROTOCOLADOS', ufi, anoi, mesi) # [PTITES-1637] #
    dir_regerados = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'REGERADOS', ufi, anoi, mesi) # [PTITES-1637] #
  
    
  
    
  
    
    dir_dev = os.getcwd()
    
    dir_enxertados = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'ENXERTADOS',ufi,anoi,mesi) # [PTITES-1637] #
     
    # [PTITES-1637] # if (configuracoes.ambiente == 'DEV'):
    # [PTITES-1637] #     dir_enxertados = dir_dev+dir_enxertados
    
    
    
    
    
    
    
    mascara_regeradoi = "SPED_"+mesanoi+"_"+ufi+"_"+iei+"_REG*.txt"
    listadeies = ies_existentes(mascara_regeradoi,dir_regerados)
    
    for iee in listadeies:
        
        log("#")
        log("#")
        log("#")
        log("-"*100)
        log("#")
        log("# INÍCIO do processamento para a IE ", iee)
        
        mascara_regerado = "SPED_"+mesanoi+"_"+ufi+"_"+iee+"_REG*.txt"
        mascara_protocolado = "SPED_"+mesanoi+"_"+ufi+"_"+iee+"_PROT*.txt"
        
        nome_regerado = nome_arquivo(mascara_regerado,dir_regerados)
        nome_protocolado = nome_arquivo(mascara_protocolado,dir_protocolados)
         
        if ((nome_regerado == "") or (nome_protocolado == "")):
            log("-"*100)
            log("#### ERRO:    Não foi processado o ENXERTO para a dupla de arquivos:")
            log("#### Arquivo regerado    = ", nome_regerado)
            log("#### Arquivo protocolado = ", nome_protocolado)
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
               
                
               
                
               
                
               
                
               
               
                
            nome_enxertado = os.path.join(dir_enxertados, "SPED_"+mesanoi+"_"+ufi+"_"+iee+"_ENX"+versao_enxertado+ ".txt")
        
            if not os.path.isdir(dir_enxertados) :
                os.makedirs(dir_enxertados)

            log("#")
            log("-"*100)
            log("#  Arquivos a serem processados:")
            log("#     Arquivo protocolado = ", nome_protocolado)
            log("#     Arquivo regerado    = ", nome_regerado)
            log("#     Arquivo enxertado   = ", nome_enxertado)
            log("-"*100)
            
            if (processaDiretorio(nome_protocolado, nome_regerado, nome_enxertado, dir_protocolados, dir_regerados, dir_enxertados) == False):
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

def parametros():
    '''
    Principal
    Não recebe parametros
    Retorna os parametros recebidos na ordem (ufi,mesanoi,mesi,anoi,iei)
    ou false para o parametro inválido.
    '''
    global ret
    ufi = ""
    mesanoi = ""
    iei = "*"
    mesi = ""
    anoi = "" 
    ret = 0
    
#### Recebe, verifica e formata os argumentos de entrada.  
    if (len(sys.argv) >= 3 ): 
        ufi = sys.argv[1].upper()
    if (len(sys.argv) >= 3 
        and validauf(ufi)
        and len(sys.argv[2])==6  
        and int(sys.argv[2][:2])>0 
        and int(sys.argv[2][:2])<13
        and int(sys.argv[2][2:])<=datetime.datetime.now().year
        and int(sys.argv[2][2:])>(datetime.datetime.now().year)-50
        ):
     
        mesanoi = sys.argv[2].upper()

        if len(sys.argv) > 3:
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
        log('####      .%s%s <UF> <MMAAAA> [IE] '%(SD,sys.argv[0]))
        log("#### ")
        log('#### Onde')
        log('####      <UF> = estado. Ex: SP')
        log('####      <MMAAAA> = mês e ano. Ex: Para junho de 2020 informe 062020')
        log('####      [IE] =     Inscição Estadual. É opcional, pode ou não ser informado.')
        log('####                 caso não informado, será processado para todas IEs do estado <UF> informado.')
        log("#### ")
        log('#### Portanto, se o estado = SP, o mes = 06 e o ano = 2020, e deseja todas IEs,  o comando correto deve ser :')  
        log('####      .%s%s SP 062020'%(SD,sys.argv[0]))  
        log("#### ")
        log('#### Se desejar processar apenas para IE 108383949112, o comando deve ser:')  
        log('####      .%s%s SP 062020 108383949112'%(SD,sys.argv[0]))  
        log("#### ")
        log('#### ')
        log("-" * 100)
        log("")
        log("Retorno = 99") 
        ret = 99
        return(False,False,False,False,False)

    mesanoi = sys.argv[2].upper()
    mesi = sys.argv[2][:2].upper()
    anoi = sys.argv[2][2:].upper()
    
    return(ufi,mesanoi,mesi,anoi,iei)


def gravar2(arquivo, dado, contador):
    contador = contador + 1
    arquivo.write(dado)
    return (contador)

def gravar(arquivo, dado, contador):
    contador = contador + 1
    chave = dado.split('|')[1]
    vi = chave[0]
    chaveti = vi + '990'

    #total por inicial
    if (chaveti in toti):
        toti[chaveti] = toti[chaveti] + 1
    else:
        toti[chaveti] = 1

    if ((chave[1:] == '990' )):
 #       log("dado = ", dado)
 #       log("chaveti = ", chaveti)
 #       log("toti[chaveti] = ", toti[chaveti])
        dado = '|'+ str(chaveti) + '|'+ str(toti[chaveti]) + '|\n'
 #       log("dado = ", dado)

    #total por valor
    chavetv = '|9900|' + str(chave) + '|'
    if (chavetv in totv):
        totv[chavetv] = totv[chavetv] + 1
    else:
        totv[chavetv] = 1

    #grava no destino
    arquivo.write(dado)
    return (contador)


def processaDiretorio(nome_protocolado, nome_regerado, nome_enxertado, path_protocolados, path_regerados, path_enxertados) :
    global relatorio_erros 
    
    path_nomeArquivoProtocolado = nome_protocolado
    path_nomeArquivoRegerado = nome_regerado
    path_nomeArquivoEnxertado = nome_enxertado

    encP = encodingDoArquivo(path_nomeArquivoProtocolado)
    arquivoP = open(path_nomeArquivoProtocolado, 'r', encoding=encP, errors='ignore')
    reg_inicialP = False
    num_linhaP = 0

#ler até encontrar o primeiro registro....
    for linhaP in arquivoP:
        num_linhaP = num_linhaP + 1
        if linhaP.startswith('|0000|') :
            reg_inicialP = linhaP[:]
            break


    encR = encodingDoArquivo(path_nomeArquivoRegerado)
    arquivoR = open(path_nomeArquivoRegerado, 'r', encoding=encR, errors='ignore')
    reg_inicialR = False
    num_linhaR = 0

#ler até encontrar o primeiro registro....
    for linhaR in arquivoR:
        num_linhaR = num_linhaR + 1
        if linhaR.startswith('|0000|') :
            reg_inicialR = linhaR[:]
            break


#Verifica se os arquivos são do mesmo periodo.
    if (linhaP.split('|')[4] != linhaR.split('|')[4]) or (linhaP.split('|')[5] != linhaR.split('|')[5]) :
                log("#"*80)
                log('### ERRO - Arquivos com periodos de dados diferentes ... Verifique !!')
                log("#"*80)
                return False
            
#Cria o arquivo enxertado 
    arquivoE = open(path_nomeArquivoEnxertado, 'w')
    numlinE = 0
#gravar linha |0000| no Enxertado
    numlinE = gravar(arquivoE, linhaP[:], numlinE)


    """
    1 - 
    Complementar os registros |0150| e |0200| que estejam no arquivo regerado e não estejam no protocolado –  ler todos os registros 0150 e 0200 do             regerado e incluir no arquivo protocolado os não encontrados (que existam no regerado e não existam no protocolado);
    
    2 -
    Substituir o registro |0400|;
    
    3 - 
    Substituir do arquivo protocolado todas as linhas contidas entre os registros |C001| e |D990|;
    
    4 - 
    Recontar os registros: |0990|, |C990|, |D990|    |9900|0150|, |9900|0200|, |9900|0400|, |9900|C100|, |9900|C101|, |9900|C110|, |9900|C170|, |9900|C190|, |9900|C500|, |9900|C590|, |9900|C990|, |9900|D001|, |9900|D100|, |9900|D190|, |9900|D500|, |9900|D590|, |9900|D695|, |9900|D696|, |9900|D697|, |9900|9900|, |9990|, |0990| e |9999|.
    """
#gravar protocolado no ArquivoE até encontrar 0150 ou maior
    for linhaP in arquivoP:
        num_linhaP = num_linhaP + 1
        if (linhaP.split('|')[1] < '0150') :
            numlinE = gravar(arquivoE, linhaP[:], numlinE)
        else:
            break
        #endif
    #endfor

#posiciona regerado na primeira linha >=150
    for linhaR in arquivoR:
        num_linhaR = num_linhaR + 1
        if (linhaR.split('|')[1] < '0150') :
            continue
        else:
            break
        #endif
    #endfor

#gravar linha >=150 e <=200 no enxertado, pega do protocolado. Se a do regerado não existir, pega do regerado.
    lerP = False
    lerR = False
    filho = False
    ler = True
    valorP = linhaP.split('|')[1]
    valorR = linhaR.split('|')[1]

    while (ler == True):

        if(lerP == True):
            for linhaP in arquivoP:
                valorP = linhaP.split('|')[1]

                if ((valorP in filhos0150) or (valorP in filhos0200) ):
                    numlinE = gravar(arquivoE, linhaP[:], numlinE)
                    continue

                if (valorP  > '0200') :
                    lerP = False
                    break

        if(lerR == True):
            for linhaR in arquivoR:
                valorR = linhaR.split('|')[1]

                if ((valorR in filhos0150) or (valorR in filhos0200 )):
                    if (lerP == False):
                        numlinE = gravar(arquivoE, linhaR[:], numlinE)
                    continue

                if ((valorR  >= '0150' and valorR  <= '0200') or valorR > '0209') :
                    lerR = False
                    break

        lerP = False
        lerR = False

        if (valorP <= '0200' or valorR <= '0200') :

            if (linhaP[:] < linhaR[:]) :
                numlinE = gravar(arquivoE, linhaP[:], numlinE)
                lerP = True
                lerR = False
                #log(numlinE,"linhaP[:] < linhaR[:]")


            if (linhaP[:] > linhaR[:]) :
                numlinE = gravar(arquivoE, linhaR[:], numlinE)
                lerP = False
                lerR = True
                #log(numlinE,"linhaR[:] < linhaP[:]")


            if (linhaP[:] == linhaR[:]) :
                numlinE = gravar(arquivoE, linhaP[:], numlinE)
                lerP = True
                lerR = True
                #log(numlinE,"linhaR[:] == linhaP[:]")
        else:
            lerP = False
            lerR = False
            ler = False
            break
    #endwhile

# Grava do protocolado até chegar no 0400
    while valorP < '0400':
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break



    """
    2 -
    Substituir o registro |0400|;
    
    3 - 
    Substituir do arquivo protocolado todas as linhas contidas entre os registros |C001| e |D990|;
    
    4 - 
    Recontar os registros: |0990|, |C990|, |D990|    |9900|0150|, |9900|0200|, |9900|0400|, |9900|C100|, |9900|C101|, |9900|C110|, |9900|C170|, |9900|C190|, |9900|C500|, |9900|C590|, |9900|C990|, |9900|D001|, |9900|D100|, |9900|D190|, |9900|D500|, |9900|D590|, |9900|D695|, |9900|D696|, |9900|D697|, |9900|9900|, |9990|, |0990| e |9999|.
    """

#posiciona o protocolado no > 0400
    while valorP <= '0400':

        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break

#posiciona o regerado no 0400
    while valorR < '0400':
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break

#grava os 0400 do regerado
    while valorR == '0400':
        numlinE = gravar(arquivoE, linhaR[:], numlinE)
        for linhaR in arquivoR:
            valorR = linhaR.split('|')[1]
            break




    """
    3 - 
    Substituir do arquivo protocolado todas as linhas contidas entre os registros |C001| e |D990|;
    
    4 - 
    Recontar os registros: |0990|, |C990|, |D990|    |9900|0150|, |9900|0200|, |9900|0400|, |9900|C100|, |9900|C101|, |9900|C110|, |9900|C170|, |9900|C190|, |9900|C500|, |9900|C590|, |9900|C990|, |9900|D001|, |9900|D100|, |9900|D190|, |9900|D500|, |9900|D590|, |9900|D695|, |9900|D696|, |9900|D697|, |9900|9900|, |9990|, |0990| e |9999|.
    """

#grava do protocolado até o C001
    while valorP < 'C001':
        
        if valorP[0] == '9':
            break
        
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break

#posiciona o regerado no início do C001
    while valorR < 'C001':
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




    """
     4 - 
    Recontar os registros: |0990|, |C990|, |D990|    |9900|0150|, |9900|0200|, |9900|0400|, |9900|C100|, |9900|C101|, |9900|C110|, |9900|C170|, |9900|C190|, |9900|C500|, |9900|C590|, |9900|C990|, |9900|D001|, |9900|D100|, |9900|D190|, |9900|D500|, |9900|D590|, |9900|D695|, |9900|D696|, |9900|D697|, |9900|9900|, |9990|, |0990| e |9999|.
    """

#posiciona o protocolado no > D990
    while valorP <= 'D990':
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break

#grava do protocolado até o |9900|
    toti['9990'] = 0

    while valorP != '9900':
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break

    """
     4 - 
    Recontar os registros: |0990|, |C990|, |D990|    |9900|0150|, |9900|0200|, |9900|0400|, |9900|C100|, |9900|C101|, |9900|C110|, |9900|C170|, |9900|C190|, |9900|C500|, |9900|C590|, |9900|C990|, |9900|D001|, |9900|D100|, |9900|D190|, |9900|D500|, |9900|D590|, |9900|D695|, |9900|D696|, |9900|D697|, |9900|9900|, |9990|, |0990| e |9999|.
    """
    qtd9900 =  0
    for soma in totv:
        qtd9900 = qtd9900 + 1
        linhaf = soma + str(totv[soma]) + '|\n'
#        log("linha = ", linhaf )
        numlinE = gravar2(arquivoE, linhaf, numlinE)


    linhaf = '|9900|9990|1|\n'
    qtd9900 = qtd9900 + 1
    numlinE = gravar2(arquivoE, linhaf, numlinE)
    
    linhaf = '|9900|9999|1|\n'
    qtd9900 = qtd9900 + 1
    numlinE = gravar2(arquivoE, linhaf, numlinE)

    linhaf = '|9900|9900|'+ str(qtd9900 + 1 )   +'|\n'
    numlinE = gravar2(arquivoE, linhaf, numlinE)

    linhaf = '|9990|' + str(qtd9900 + 3 + toti['9990'] ) + '|\n'
    numlinE = gravar2(arquivoE, linhaf, numlinE)

    linhaf = '|9999|' + str(numlinE + 1) + '|\n'
    numlinE = gravar(arquivoE, linhaf, numlinE)

    arquivoP.close()
    arquivoR.close()
    arquivoE.close()

    return True

if __name__ == "__main__":
    global ret
    
    log('#'*100)
    log("# ")  
    log("# - INICIO - ENXERTO_SPED_PE")
    log("# ")
    log('#'*100)
    ret = 0
    retorno = parametros()
    
    ufi = retorno[0]
    if (retorno[0] != False):
        ret     = 0
        ufi     = retorno[0]
        mesanoi = retorno[1]
        mesi    = retorno[2]
        anoi    = retorno[3]
        iei     = retorno[4]
        log("-"*100)
        log("# Processando ENXERTO SPED para os seguintes parâmetros:")
        log("#    UF  = ",retorno[0])
        log("#    MÊS = ",retorno[2])
        log("#    ANO = ",retorno[3])
        log("#    IE  = ",retorno[4])
        log("-"*100)
        ret = processar(ufi,mesanoi,mesi,anoi,iei)
    else:
        ret = 99
        
    log('#'*100)
    log("# ")  
    log("# - FIM - ENXERTO_SPED_PE")
    log("# ")
    log("#"*100)

    log("Codigo de saida = ",ret)
    sys.exit(ret)
 
 
