#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-

"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: gerar_SPED.py
  CRIACAO ..: 17/11/2020
  AUTOR ....:  Airton Borges da Silva Filho - Kyros Consultoria 
  DESCRICAO : Agendar e acompanhar a execução do SPED - GERAR LIVROS FISCAIS P9
                

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

----------------------------------------------------------------------------------------------
"""
import sys
import os
SD = '/' if os.name == 'posix' else '\\'
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV')[0], 'DEV')
sys.path.append(dir_base)
import configuracoes
import comum
from comum import log
import sql


import datetime
import time
import shutil
import re
from pathlib import Path
global variaveis
global db 

#DEBUG = True
DEBUG = False

#log.gerar_log_em_arquivo = True

db =  ""
 
ret = 0

comum.carregaConfiguracoes(configuracoes)
cursor=sql.geraCnxBD(configuracoes)

def dtf():
    return (datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'))

def ultimodia(ano,mes):
   return(31 if mes == 12 else datetime.date.fromordinal((datetime.date(ano,mes+1,1)).toordinal()-1).day)

def processar():
    global variaveis
    global db
    mesi = ""
    anoi = ""
    ret = 0

    if (len(sys.argv) != 4 ): 
        ret = 1 # PARAMETROS INVALIDOS

#### Recebe, verifica e formata os argumentos de entrada.
    if (ret == 0 # <<20210702>>

        and len(sys.argv[1])==6  
        and int(sys.argv[1][:2])>0 
        and int(sys.argv[1][:2])<13
        and int(sys.argv[1][2:])<=datetime.datetime.now().year
        and int(sys.argv[1][2:])>(datetime.datetime.now().year)-50

        and len(sys.argv[2])==6  
        and int(sys.argv[2][:2])>0 
        and int(sys.argv[2][:2])<13
        and int(sys.argv[2][2:])<=datetime.datetime.now().year
        and int(sys.argv[2][2:])>(datetime.datetime.now().year)-50
        
        and len(sys.argv[3]) > 0 
        ):
     
        
        mesi  = sys.argv[1][:2].upper()
        anoi  = sys.argv[1][2:].upper()
        datai = "01/"+mesi+"/"+anoi

        mesf  = sys.argv[2][:2].upper()
        anof  = sys.argv[2][2:].upper()
        dataf = str(ultimodia(int(anof),int(mesf)))+"/"+str(mesf)+"/"+str(anof)
       
        pdiai = str(anoi)+str(mesi)+"01"
        udiai = str(anof)+str(mesf)+str(ultimodia(int(anof),int(mesf)))
               
        iei   = sys.argv[3]
        iei = re.sub('[^0-9]','',iei)
        if ( (iei == "") or (iei == "''") or (iei == '""') or (int("0"+iei) == 0)):
            iei = "*"

        if(DEBUG):
            print()
            print()
            print(mesi)
            print(anoi)
            print(datai)
            print()
            print(mesf)
            print(anof)
            print(dataf)
            print()
            print(iei)
            print()
            print(pdiai)
            print(udiai)
            print()
            print()

    else :
        # comum.imprimeHelp() usar este ou o formato abaixo
        log("-" * 100)
        log("#### ")
        log('#### ERRO - Erro nos parametros do script.')
        log("#### ")
        log('#### Exemplo de como deve ser :')
        log('####  %s <MMAAAA-inicio> <MMAAAA-fim> <IE>'%(sys.argv[0]))
        log("#### ")
        log('#### Onde')
        log('####  <MMAAAA-inicio>        = mês e ano inicial. Ex: Para junho de 2020 informe 062020')
        log('####  <MMAAAA-fim>           = mês e ano final . Ex: Para julho de 2020 informe 072020')
        log('####  <IE>                   = Inscrição Estadual')
        log("#### ")
        log("-" * 100)
        log("")
        log("Retorno = 99") 
        ret = 99
        return(99)

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

#### Seleciona um usuário livre ou termina com erro após TRYX tentativas 
    uemuso = 1
    qtdver = 0
    ES = int(configuracoes.waitsegs)
    TX = int(configuracoes.tryx)
    
    if (DEBUG):
        print()
        print("Espera entre tentativas em segundos, ES = ", ES)
        print("Quantidade de tentativas,            TX = ", TX)
        print()
    
    while (qtdver < TX and uemuso > 0) :
        qtdver = qtdver + 1
        log("-"*100)
        log("##", qtdver,"/ ", TX,  " - procurando por um usuário livre...")
       
        for u in usuario:
            log("## Verificando o usuário", u, "...")
            uemuso = usuario_livre(u)
            if (uemuso == 0):
                break
            log("#### ATENÇÃO #### ==>  USUÁRIO " , u, " OCUPADO.")                 
 
        if (uemuso > 0 and qtdver < TX):
            log("#### ATENÇÃO #### ==> NA TENTATIVA TENTATIVA ", qtdver,"/",TX,",   TODOS OS USUÁRIOS ESTAVAM OCUPADOS. VERIFICANDO NOVAMENTE EM ", ES, " SEGUNDOS")  
               
            time.sleep(ES)
            
    if (uemuso > 0 and qtdver > (TX - 1)):
        log("#### ERR0 #### ==>  TEMPO DE ESPERA PARA LIBERAÇÃO DE USUÁRIO ESGOTADO. TODOS USUÁRIOS ESTAVAM OCUPADOS EM ", TX ," TENTATIVAS DE ", ES, " EM ", ES," SEGUNDOS")
        return(99)
        
    usrlivre = u
    
#### Busca novo sequencial de execução no Banco de dados ( sequence )
    id_num_seq = 0
    #print("id_num_seq",id_num_seq )    
    id_num_seq = prox_id_num_seq()
    #print("id_num_seq", id_num_seq)    
        
 #### Define o ID_AGENDA ( sequence )
    id_agenda = 0
    #print("id_agenda", id_agenda)    
    id_agenda = prox_id_agenda()   
    #print("id_agenda",id_agenda )    

    if (DEBUG):
        print()
        print("Sequencial de execucao = ", id_num_seq)
        print("Id_Agenda              = ", id_agenda)
        print()

    
    dhip = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')      
    amdhm = datetime.datetime.now().strftime('%Y%m%d%H%M')
    dip = datetime.datetime.now().strftime('%d/%m/%Y')
    
    if ( DEBUG ):
        log(" = ", )
        log("###### debug ######")
        log("db                  = ", db)
        for u in usuario:
            log("Usuário             = ", u) 
        log("ID Agenda           = ", id_agenda)
        log("Usuário livre       = ", usrlivre)
        log("IE                  = ", iei)
        log("data_início         = ", datai)
        log("data_fim            = ", dataf)
        log("data_inicio_invert  = ", pdiai)
        log("data_fim_invert     = ", udiai)
        log("Data_inicio_Proc    = ", dip)
        log("Data_Hora_Inic_Proc = ", dhip)
        log("AAAAMMDDHHMM        = ", amdhm)
        log("###### debug ######")
        log(" = ", )

    agendamento = insere_agen_param_proc(id_agenda,id_num_seq,usrlivre,iei,datai,dataf,pdiai,udiai,amdhm,dip,dhip) 
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
    statusag = status_agenda(id_num_seq)
    
    while ( statusag == "A" or statusag == "E" ):
        log("##",dtf(), "- STATUS DO AGENDAMENTO = ",statusag)
        time.sleep(60)
        statusag = status_agenda(id_num_seq)
    log("##",dtf(), "- STATUS DO AGENDAMENTO = ",statusag)
 
    return(0)
     
def insere_agen_param_proc(idag,idseq,ul,ie,di,df,dii,dfi,amdh,dip,dhip):
    td_di = "to_date('" + di + "','DD/MM/RRRR')"
    td_df = "to_date('" + df + "','DD/MM/RRRR')"
    c=chr(27)
    q="TO_CLOB(q'[empresa.resumo_fiscal=TBRA"           + c
    q=q+"filial.resumo_fiscal="                         + c
    q=q+"ie_filial.resumo_fiscal="              + ie    + c
    q=q+"p_dat_ini.resumo_fiscal="              + td_di + c
    q=q+"p_dat_fim.resumo_fiscal="              + td_df + c
    q=q+"$$vg_usr_current="                     +  ul   + c
    q=q+"entradas.resumo_fiscal=T"                      + c
    q=q+"saidas.resumo_fiscal=T"                        + c
    q=q+"em.resumo_fiscal=T"                            + c
    q=q+"et.resumo_fiscal="                             + c
    q=q+"sm.resumo_fiscal=T"                            + c
    q=q+"se.resumo_fiscal=F"                            + c
    q=q+"st.resumo_fiscal=T"                            + c
    q=q+"sc.resumo_fiscal=F"                            + c
    q=q+"mat_cen.resumo_fiscal="                        + c
    q=q+"sa.resumo_fiscal="                             + c
    q=q+"ind_valores_decl.resumo_fiscal="               + c
    q=q+"apuracao.resumo_fiscal="                       + c
    q=q+"opcao.resumo_fiscal=09']')"  
    
   
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
                                ,'CALC'
                                ,'A'
                                ,null
                                ,%s
                                )

            """%(idag,idseq,ul,ie,td_di,td_df,q)
    
 
    
          
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
            """%(idseq,vdip,vdhip)
            
####DEBUG#### 
    if (DEBUG):
        log(" = ", )
        log("query1 = ",query1 )
        log(" = ", )
        log("query2 = ",query2 )
        log(" = ", )
        input("CONTINUA?")
    try:           
        cursor.executa(query1)
        cursor.executa(query2)
#        connection.rollback()
        cursor.commit()
        log('INSERT REALIZADO NA openrisow.AGEN_PARAM: ', q)
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

    cursor.executa(query)
    result = cursor.fetchone()

    if result:
        return result[0]
    return None

def prox_id_num_seq():
    query = """
        select 
            NVL(MAX(agen_num), 0) + 1 PROXIMO_AGEN_NUM
        from 
            openrisow.AGEN_PROC
        """
    cursor.executa(query)
    result = cursor.fetchone()
    for v in result:
        break
    return (v)
 
def prox_id_agenda():
    query = """
        select 
            NVL(max(AGEN_PROC_ID), 0) + 1 PROXIMO_AGEN_PROC_ID
        from 
            openrisow.AGEN_PARAM
        where 
            AGEN_TP_PROC = 'CALC'
        """

    cursor.executa(query)
    result = cursor.fetchone()
    for v in result:
        break
    return (v)

def usuario_livre(usr):

    query1 = """
        select
            count(1)
        from
            OPENRISOW.log_control
        where
            1=1
            and MOD_COD = 'CALC_9X'
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
            and agen_tp_proc = 'CALC'
            and agen_USR = '%s'
        """%(usr)


    cursor.executa(query1)
    result1 = cursor.fetchone()
    
    cursor.executa(query2)
    result2 = cursor.fetchone()
    
    for r1 in result1:
        break
    for r2 in result2:
        break    

    return (0 +(0 if(r1 == None) else r1) + (0 if(r2 == None) else r2))

if __name__ == "__main__":
    ret = processar()
    log("## Código de execução = ", ret)
    log("#### ",dtf(), " FIM DO AGENDAMENTO SPED GERAR LIVROS FISCAIS P9 ####")
    if ret > 0 :
        log("ERRO - Verifique as mensagens anteriores")
    sys.exit(ret)
