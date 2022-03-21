#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-

"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: gerar_SPED.py
  CRIACAO ..: 17/11/2020
  AUTOR ....:  Airton Borges da Silva Filho - Kyros Consultoria 
  DESCRICAO : Agendar e acompanhar a execução do SPED
                

----------------------------------------------------------------------------------------------
  HISTORICO : 
    07/12/2020 - Finalizado.
    09/12/2020 - Airton - Alterados parametros da query blocoB.dummy e blocoC.dummy, ambos de T para F.
        - incluido a verificacao do IE se é aspas e a biblioteca re
    02/07/2021 - eduardof@kyros.com.br 
                 Colocar a opcao por bloco1 por parametrizacao
                 tag : <<20210702>>    

    04/08/2021 - Airton - 
            Alterado para contemplação do bloco C e “Cálculo dos valores agregados”.
            Utilizar os seguintes parâmetros de entrada:
                <UF>: Obrigatório, possuindo a sigla da UF, exemplo: SP.
                SANO> = Obrigatório, possuindo o mês seguido do ano a ser processado, exemplo: 032020
                <IE>: Opcional, exemplo: 108383949112
                <Bloco1>: Opcional, exemplo S [default] ou N 
                <Bloco_C>: Opcional, exemplo S ou N [default]
                ilizar_TAB_INVA>: Opcional, exemplo S ou N [default]
                oco_C_Convencional>: Opcional, exemplo S ou N [default]
    31/08/2021 - Arthur - PTITES-136
               - Ajuste do script para o novo Padrão.            
    * 22/02/2022 - Eduardo da Silva Ferreira - Kyros Tecnologia
                 - [PTITES-1636] Padrão de diretórios do SPARTA                 

----------------------------------------------------------------------------------------------
"""
import sys
import os
SD = '/' if os.name == 'posix' else '\\'
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV')[0], 'DEV')
sys.path.append(dir_base)
import configuracoes
import comum
import sql


import datetime
import time
import shutil
import re
from pathlib import Path
global variaveis
global db 
global arquivo_destino

DEBUG = False

#log.gerar_log_em_arquivo = True

db =  ""
 
ret = 0
nome_relatorio = "" 
dir_destino = "" 
dir_base = "" 
arquivo_destino = ""

comum.carregaConfiguracoes(configuracoes)

def validauf(uf):
    return(True if (uf.upper() in ('AC','AL','AM','AP','BA','CE','DF','ES','GO','MA','MG','MS','MT','PA','PB','PE','PI','PR','RJ','RN','RO','RR','RS','SC','SE','SP','TO')) else False)
          
def dtf():
    return (datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'))

def ultimodia(ano,mes):
   return(31 if mes == 12 else datetime.date.fromordinal((datetime.date(ano,mes+1,1)).toordinal()-1).day)


def proximo_arquivo(mascara,diretorio):
    qdade = 0
    nomearq = "" 
    directory = Path(diretorio)
    files = directory.glob(mascara)
    sorted_files = sorted(files, reverse=False)    
   
    nomearq = mascara.replace("*", "V000")
     
    if sorted_files:
        for f in sorted_files:
            qdade = qdade + 1
            nomearq = f
    proximo = '{:03d}'.format(int((str(nomearq).split(".")[0]).split("_")[-1][1:]) + 1)
    nomearq = mascara.replace("*", "V"+proximo)
    return(nomearq)

def processar():
    global variaveis
    global db
    global arquivo_destino
    ufi = ""
    mesanoi = ""
    mesi = ""
    anoi = ""
    ret = 0
    # INICIO <<20210702>>
    l_bloco1        = "T"
    l_blococ        = "F"
    l_tabinva       = "F"
    l_blococconv    = "F"

    if (len(sys.argv) < 4 ): 
        ret = 1 # PARAMETROS INVALIDOS

    if (len(sys.argv) >= 4 ): 
        ufi = sys.argv[1].upper()

    if (len(sys.argv) >= 5 ): 
        l_bloco1     = sys.argv[4].upper().strip()
        l_blococ     = sys.argv[5].upper().strip()
        l_tabinva    = sys.argv[6].upper().strip()
        l_blococconv = sys.argv[7].upper().strip()

        

        if ( l_bloco1 == "N"
            or l_bloco1 == "F"):
            l_bloco1 = "F"
        else:
            l_bloco1 = "T"

        if ( l_blococ != "S"
            or l_blococ == "F"):
            l_blococ = "F"
        else:
            l_blococ = "T"

        if ( l_tabinva != "S"
            or l_tabinva == "F"):
            l_tabinva = "F"
        else:
            l_tabinva = "T"

        if ( l_blococconv != "S" or l_blococconv == "F"):
            l_blococconv = "F"
        else:
            l_blococconv = "T"

    else:
        ret = 1 # PARAMETROS INVALIDOS    

#### Recebe, verifica e formata os argumentos de entrada.
    if (ret == 0 # <<20210702>>
        and validauf(ufi)
        and len(sys.argv[2])==6  
        and int(sys.argv[2][:2])>0 
        and int(sys.argv[2][:2])<13
        and int(sys.argv[2][2:])<=datetime.datetime.now().year
        and int(sys.argv[2][2:])>(datetime.datetime.now().year)-50
        ):
     
        mesanoi = sys.argv[2].upper()
        mesi  = sys.argv[2][:2].upper()
        anoi  = sys.argv[2][2:].upper()
        datai = "01/"+mesi+"/"+anoi
        dataf = str(ultimodia(int(anoi),int(mesi)))+"/"+str(mesi)+"/"+str(anoi)
        iei   = sys.argv[3]
        iei = re.sub('[^0-9]','',iei)
        if ( (iei == "") or (iei == "''") or (iei == '""') or (int("0"+iei) == 0)):
            iei = "*"

    else :
        # comum.imprimeHelp() usar este ou o formato abaixo
        log("-" * 100)
        log("#### ")
        log('#### ERRO - Erro nos parametros do script.')
        log("#### ")
        log('#### Exemplo de como deve ser :')
        log('####      %s <UF> <MMAAAA> <IE> <BLOCO_1> <BLOCO_C> <UTILIZAR_TAB_INVA> <BLOCO_C_CONVENCIONAL>'%(sys.argv[0]))
        log("#### ")
        log('#### Onde')
        log('####      <UF>                   = estado. Ex: SP')
        log('####      <MMAAAA>               = mês e ano. Ex: Para junho de 2020 informe 062020')
        log('####      <IE>                   = Inscrição Estadual')
        log('####      <BLOCO1>               = (S - SIM / N - NAO)')
        log('####      <BLOCO_C>              = (S - SIM / N - NAO)')
        log('####      <Utilizar_TAB_INVA>    = (S - SIM / N - NAO)')
        log('####      <Bloco_C_Convencional> = (S - SIM / N - NAO)')
        log("#### ")
        log('#### Portanto, se o estado = SP, o mes = 06 e o ano = 2020, e IE = 108383949112 o comando correto deve ser :')  
        log('####      %s SP 062020 108383949112 S N N N'%(sys.argv[0]))  
        log("#### ")
        log('#### ')
        log("-" * 100)
        log("")
        log("Retorno = 99") 
        ret = 99
        return(99)

#### Monta caminho e nome do destino
    dir_base = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'REGERADOS') # [PTITES-1636] # configuracoes.dir_base
    dir_destino = os.path.join(dir_base, ufi, anoi, mesi)
    
#### Monta caminho e nome do origem (gerado )
    dir_origem = os.path.join(SD + 'arquivos' + SD, 'SPED_FISCAL') # dir_origem = configuracoes.dir_origem # os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL') # [PTITES-1636] #   

#### Se a pasta do relatório não existir, cria
    if not os.path.isdir(dir_destino) :
        os.makedirs(dir_destino)    

#### Monta o nome do próximo arquivo (VERSÃO V001 ou última + 1 )    
    maskfile =  "SPED_" + mesanoi + "_" + ufi + "_"+ iei +"_REG_*.txt"
    arquivo_destino = proximo_arquivo(maskfile,dir_destino)
    arquivo_destino = os.path.join(dir_destino,arquivo_destino)
    arquivo = open(arquivo_destino, 'w')
    arquivo.close()
     
#### Carrega todos os usuários      
    usuario = []
    i = 1
    while  (getattr(configuracoes, 'usuario_%s'%(i), False)):
        usuario.append(getattr(configuracoes,'usuario_%s'%(i)))
        i= i+1
        
    if (i == 1):
        log("#### ERR0 #### ==>  NÃO EXISTEM USUARIOS NO ARQUIVO ", sys.argv[0],".cfg")
        return(99)
#### Carrega o nome do banco 
    db = configuracoes.banco if configuracoes.banco else db
    
#### define o dia inicial e final no formato invertido AAAAMMDD     
    pdiai = str(anoi)+str(mesi)+"01"
    udiai = str(anoi)+str(mesi)+str(ultimodia(int(anoi),int(mesi)))
    

#### Seleciona um usuário livre ou termina com erro após 12 tentativas 
    uemuso = 1
    qtdver = 0
    
    while (qtdver < 12 and uemuso > 0) :
        qtdver = qtdver + 1
        log("-"*100)
        log("##", qtdver,"/ 12 -", "procurando por um usuário livre...")
       
        for u in usuario:
            log("## Verificando o usuário", u, "...")
            uemuso = usuario_livre(u)
            if (uemuso == 0):
                break
            log("#### ATENÇÃO #### ==>  USUÁRIO " , u, " OCUPADO.")                 
 
        if (uemuso > 0 and qtdver < 12):
            log("#### ATENÇÃO #### ==> NA TENTATIVA TENTATIVA ", qtdver,"/12 TODOS OS USUÁRIOS ESTAVAM OCUPADOS, VERIFICANDO NOVAMENTE EM 1 MINUTO...")  
               
            time.sleep(60)
            
    if (uemuso > 0 and qtdver > 11):
        log("#### ERR0 #### ==>  TEMPO DE ESPERA PARA LIBERAÇÃO DE USUÁRIO ESGOTADO. TODOS USUÁRIOS ESTAVAM OCUPADOS EM 12 TENTATIVAS EM 12 MINUTOS")
        return(99)
        
    usrlivre = u
            
     
#### Define a versão, iniciando na 002
    versaocz = ""
    versaosz = ""   
    n=2   
    i="002"

    while  (getattr(configuracoes,'versao_%s'%(i), False)): 
        dvi= getattr(configuracoes,'versao_%s'%(i))[:8]
        dvf= getattr(configuracoes,'versao_%s'%(i))[8:] 
         
        if ((pdiai >= dvi) and (pdiai <= dvf)):
            versaocz = i
            versaosz = str(n)
            break

        i='{:03d}'.format(n+1)  
        n=n+1
        
    if (versaosz == "" or versaocz == ""):
        log("#### ERR0 #### ==>  NÃO ENCONTRADO A VERSÃO NO ARQUIVO Agenda_SPED.cfg REFERENTE A DATA INFORMADA.")
        return(99)
   
#### Define o CNPJ

    indicecnpj = ufi+iei 
    cnpj = getattr(configuracoes, indicecnpj, None)
        
    if (cnpj == None):
        log("#### ERR0 #### ==>  NÃO ENCONTRADO O CNPJ PARA ESTADO = ", ufi, " E IE = ", iei)
        return(99)

 #### Define o ID_AGENDA
    id_agenda = prox_id_agenda()      
 
    dhip = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')      
    amdhm = datetime.datetime.now().strftime('%Y%m%d%H%M')
    dip = datetime.datetime.now().strftime('%d/%m/%Y')

    if ( DEBUG ):
        log(" = ", )
        log("###### debug ######")
        log("dir_base            = ", dir_base)    
        log("dir_destino         = ", dir_destino)
        log("maskfile            = ", maskfile)
        log("arquivo_destino     = ", arquivo_destino)
        log("db                  = ", db)
        for u in usuario:
            log("Usuário             = ", u) 
        log("Versão 1            = ", versaosz)
        log("ID Agenda           = ", id_agenda)
        log("Usuário livre       = ", usrlivre)
        log("IE                  = ", iei)
        log("data_início         = ", datai)
        log("data_fim            = ", dataf)
        log("CNPJ                = ", cnpj)
        log("data_inicio_invert  = ", pdiai)
        log("data_fim_invert     = ", udiai)
        log("Versão 2            = ", versaocz)
        log("MêsAno              = ", mesanoi)
        log("Data_inicio_Proc    = ", dip)
        log("Data_Hora_Inic_Proc = ", dhip)
        log("AAAAMMDDHHMM        = ", amdhm)
        log("###### debug ######")
        log(" = ", )


### AGENDA A EXECUÇÃO...
    arq_origem = "SPED_"+mesanoi+"_"+ufi+"_"+iei+"_REG_"+usrlivre+"_"+amdhm+".txt"
    outros_origem = "SPED_"+mesanoi+"_"+ufi+"_"+iei+"_REG_"+usrlivre+"_"+amdhm+".*"
 
    maskfile =  "SPED_" + mesanoi + "_" + ufi + "_"+ iei +"_REG_*.txt"
#    arquivo_destino = proximo_arquivo(maskfile,dir_destino)
    ao = os.path.join(dir_origem,arq_origem)
    ad = os.path.join(dir_destino,arquivo_destino)
    
    oao = os.path.join(dir_origem,outros_origem)
    
    log("## AGENDANDO A GERAÇÃO DO ARQUIVO => ", ao, "...")
    log("## ARQUIVO DESTINO                => ", ad)
#    input("CONTINUA ? ")
    agendamento = insere_agen_param_proc(versaocz,id_agenda,usrlivre,iei,datai,dataf,cnpj,pdiai,udiai,versaosz,mesanoi,ufi,amdhm,dip,dhip,l_bloco1,l_blococ,l_tabinva,l_blococconv)
    if (agendamento != 0):
        log("#### ERR0 #### ==>  NÃO AGENDADO...")
        return(99)
    
    
### ACOMPHANHA O AGENDAMENTO...
    """
    A: Agendado
    E: Executando
    P: Processamento Concluído
    C: Cancelado
    """
    log("## ACOMPANHANDO O AGENDAMENTO A CADA MINUTO...")
    statusag = status_agenda(id_agenda)
    
    while ( statusag == "A" or statusag == "E" ):
        log("##",dtf(), "- STATUS DO AGENDAMENTO = ",statusag)
        time.sleep(60)
        statusag = status_agenda(id_agenda)
    log("##",dtf(), "- STATUS DO AGENDAMENTO = ",statusag)
 
### APOS A EXECUCAO...    
    if (statusag == "P") and os.path.isfile(ao):
        move_arquivo(ao,ad)
    directory = Path(dir_origem)
    maskm = outros_origem
    files = directory.glob(maskm)
    for noao in files:
        npo = os.path.join(dir_origem,noao)
        move_arquivo(npo,dir_destino)
    return(0)
     
    # move_arquivo(oao,dir_destino)
    
    # log("#### ERR0 #### ==>  RETORNO APÓS EXECUCAO AGENDADA =  ==>>", statusag , "<<==" )
    # return(99)

def move_arquivo(ao,ad):
    log("## MOVENDO O ARQUIVO ", ao, " PARA ", ad)
    shutil.move(ao, ad)
    return(0)
 
def insere_agen_param_proc(vsz,idag,ul,ie,di,df,cnpj,dii,dfi,vcz,ma,uf,amdh,dip,dhip,p_bloco1,p_blococ,p_tabinva,p_blococconv):
    c=chr(27)                                                                                                                                        
    q="TO_CLOB(q'[empresa.dummy=TBRA"                                                           +c
    q=q+"emp_desc.dummy=TBRA"                                                                   +c
    q=q+"fil_desc.dummy="                                                                       +c
    q=q+"filial.dummy="                                                                         +c
    q=q+"fili_cod_pad.dummy="                                                                   +c
    q=q+"fil_desc_pad.dummy="                                                                   +c
    q=q+"ie_filial.dummy="              + ie                                                    +c
    q=q+"fil_pad.dummy="                                                                        +c
    q=q+"cnpj.dummy="                   + cnpj                                                  +c
    q=q+"ie.dummy=T"                                                                            +c
    q=q+"data_de.dummy="                + dii                                                   +c
    q=q+"data_ate.dummy="               + dfi                                                   +c
    q=q+"versao.dummy="                 + vsz                                                   +c
    q=q+"sf_perfil.dummy=A"                                                                     +c
    q=q+"finalidade.dummy=1"                                                                    +c
    q=q+"arquivo.dummy="                + "SPED_"+ma+"_"+uf+"_"+ie+"_REG_"+ul+"_"+amdh+".txt"   +c
    q=q+"data_inv.dummy"                                                                        +c
    q=q+"bloco0.dummy=T"                                                                        +c
    q=q+"blocoB.dummy=F"                                                                        +c
    q=q+"blocoC.dummy="                 + p_blococ                                              +c
    q=q+"blocoD.dummy=T"                                                                        +c
    q=q+"blocoE.dummy=T"                                                                        +c
    q=q+"blocoG.dummy="                                                                         +c
    q=q+"blocoH.dummy=F"                                                                        +c
    q=q+"blocoK.dummy=F"                                                                        +c
    q=q+"bloco1.dummy="                 + p_bloco1                                              +c
    q=q+"ind_filial_unica.]')||TO_CLOB(q'[dummy=T"                                              +c
    q=q+"renum.dummy=T"                                                                         +c
    q=q+"ger_obs.dummy=T"                                                                       +c
    q=q+"sf_ciap.dummy=F"                                                                       +c
    q=q+"des_fret.dummy="                                                                       +c
    q=q+"des_dif.dummy=T"                                                                       +c
    q=q+"des_ress.dummy=T"                                                                      +c
    q=q+"ger_g126.dummy=F"                                                                      +c
    q=q+"ind_trans.dummy=F"                                                                     +c
    q=q+"ind_g_parcial.dummy=F"                                                                 +c
    q=q+"ckb_conv52.dummy="                                                                     +c
    q=q+"ckb_agreg.dummy="                 + p_tabinva                                          +c                                                                      +c
    q=q+"data_inve.dummy="                                                                      +c
    q=q+"motivo_inve.dummy="                                                                    +c
    q=q+"$$vg_usr_current="             + ul                                                    +c
    q=q+"$data_inv$="                                                                           +c
    q=q+"$motivo_inv$="+(c+"ind_blococ_uniface.dummy=T" if (p_blococconv == 'T') else "") +"]')"
    
    log('INSERT REALIZADO NA openrisow.AGEN_PARAM: ', q)
    
    vdi = "to_date('" + di + " 00:00:00','DD/MM/YYYY HH24:MI:SS')"
    vdf = "to_date('" + df + " 00:00:00','DD/MM/YYYY HH24:MI:SS')"
    
    vdip  = "to_date('" + dip + " 00:00:00','DD/MM/YYYY HH24:MI:SS')"
    vdhip = "to_date('" + dhip + "','DD/MM/YYYY HH24:MI:SS')"

    query1 = """ insert into 
                    openrisow.AGEN_PARAM 
                        (
                                AGEN_PROC_ID
                                ,AGEN_NUM
                                ,AGEN_NUM_SEQ
                                ,AGEN_USR
                                ,EMPRESA
                                ,FILIAL
                                ,UTILIZA_IE
                                ,IE_FILIAL
                                ,AGEN_DT_INI_GER
                                ,AGEN_DT_FIN_GER
                                ,AGEN_TP_PROC
                                ,AGEN_STATUS_PROC
                                ,AGEN_CAM_ARQ
                                ,AGEN_PARAM
                                ) 
                            values
                                (
                                '%s'
                                ,'%s'
                                ,'1'
                                ,'%s'
                                ,'TBRA'
                                ,null
                                ,'T'
                                ,'%s'
                                ,%s
                                ,%s
                                ,'SPEF'
                                ,'A'
                                ,null
                                ,%s
                                )

            """%(vcz,idag,ul,ie,vdi,vdf,q)
            
    query2 = """ insert into 
                    openrisow.AGEN_PROC
                        (AGEN_NUM
                         ,AGEN_DT
                         ,AGEN_DT_INI_PROC
                         ,AGEN_STATUS) 
                    values 
                        ('%s'
                         ,%s
                         ,%s
                         ,'A')
            """%(idag,vdip,vdhip)
            
####DEBUG#### 
    if (DEBUG):
        log(" = ", )
        log("query1 = ",query1 )
        log(" = ", )
        log("query2 = ",query2 )
        log(" = ", )
#        input("CONTINUA?")
    try:           
        cursor = sql.geraCnxBD(configuracoes)
        cursor.executa(query1)
        cursor.executa(query2)
#        connection.rollback()
        cursor.commit()
        retorno = 0
    except Exception as error:
        log("#### ERRO AO INSERIR O AGENDAMENTO NAS TABELAS openrisow.AGEN_PARAM e openrisow.AGEN_PROC")
        log(error)
        retorno = 99
        
    return(retorno)


def status_agenda(idag):

    query = """
        select 
            agen_status 
        from 
            openrisow.agen_proc
        where 
            1=1
            and agen_num = '%s'
        order by 
            agen_num desc
        """%(idag)

    cursor = sql.geraCnxBD(configuracoes)
    cursor.executa(query)
    result = cursor.fetchone()

    if result:
        return result[0]
    return None
 
def prox_id_agenda():

    query = """
        select 
            agen_num 
        from 
            openrisow.agen_proc
        where 
            1=1
            and rownum = 1
        order by 
            agen_num desc
        """
    cursor = sql.geraCnxBD(configuracoes)
    cursor.executa(query)
    result = cursor.fetchone()
    
    for v in result:
        break

    return ((0 if(v == None) else int(v)) + 1)

def usuario_livre(usr):

    query1 = """
        select
            count(1)
        from
            OPENRISOW.log_control
        where
            1=1
            and MOD_COD = 'FSPEDF00001'
            and user_id = '%s'
            and flag = 0
        """%(usr)

    query2 = """
        select
            count(1)
        from
            OPENRISOW.agen_param
        where
            1=1
            and agen_status_proc not in ('P','C')
            and agen_tp_proc = 'SPEF'
            and agen_USR = '%s'
        """%(usr)

    cursor1 = sql.geraCnxBD(configuracoes)
    cursor1.executa(query1)
    result1 = cursor1.fetchone()
    
    cursor1.executa(query2)
    result2 = cursor1.fetchone()
    
    for r1 in result1:
        break
    for r2 in result2:
        break    

    return (0 +(0 if(r1 == None) else r1) + (0 if(r2 == None) else r2))

if __name__ == "__main__":

    ret = processar()
    if (ret > 0) :
        if(arquivo_destino):
            if os.path.isfile(arquivo_destino):
                os.remove(arquivo_destino)
    log("## Código de execução = ", ret)
    log("#### ",dtf(), " FIM DO AGENDAMENTO SPED ####")
    if ret > 0 :
        log("ERRO - Verifique as mensagens anteriores")
    sys.exit(ret)