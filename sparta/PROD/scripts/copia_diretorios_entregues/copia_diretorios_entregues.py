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
import glob
import shutil

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
    
def anomesEntre(DT_MESANO_INICIO,DT_MESANO_FIM):
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
    while anomesfim >= anomesatu:
        lst_mesano.append(anomesatu[2:])
        if (int(mesatu) == 12):
            mesatu = '01'
            anoatu = str(int(anoatu)+1).zfill(4)
        else:
            mesatu = str(int(mesatu) + 1).zfill(2)
        anomesatu = str(anoatu).zfill(4) + str(mesatu).zfill(2)        
    return(lst_mesano)
    
    
def processar(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,UFI="SP") :
    ret = 0
    log("Gerando lista de todas as pastas existentes ....")
    pasta = configuracoes.dir_base_copia
    lista_datas = anomesEntre(DT_MESANO_INICIO,DT_MESANO_FIM)
   
    maskobrig = os.path.join(pasta,UFI,"??","??","TBRA",CC_FILIAL.replace("'",""),"SERIE","*","OBRIGACAO")      

    CC_SERIE = CC_SERIE.replace("'","")
    lst_serie = CC_SERIE.split(",")
    #print("lst_serie = ", lst_serie)

    lst_path_obrigacao = []    
    lst_path_temp = glob.glob(maskobrig)
    
    for dirobrig in lst_path_temp:
        log("Verificando pasta: ", dirobrig)
        caminho = dirobrig.split(dirobrig[0])
        if len(caminho) > 16 :
            anopath=caminho[11]
            mespath=caminho[12]
            filialpath=caminho[14]
            seriepath=caminho[16]
        else:     
            anopath=00
            mespath=00
            filialpath=0000
            seriepath=0000
            
            
# =============================================================================
#         
#         print("caminho          = ", caminho)
#         print("anopath          = ", anopath)
#         print("mespath          = ", mespath)
#         print("filialpath       = ", filialpath)
#         print("seriepath        = ", seriepath)
#         print("CC_FILIAL        = ", CC_FILIAL.replace("'",""))
#         print("CC_SERIE         = ", CC_SERIE.replace("'",""))
#         print("DT_MESANO_INICIO = ", DT_MESANO_INICIO)
#         print("DT_MESANO_FIM    = ", DT_MESANO_FIM)
#         
# =============================================================================
       
        if filialpath == CC_FILIAL.replace("'","") : 
            #print( "filial atende, ", filialpath)
            anomespath = anopath+mespath
            if (anomespath in lista_datas) :
                #print("AnoMes atende, ", anomespath)
                if(seriepath in lst_serie or CC_SERIE == ""):
                    #print("Id Serie atende, ", seriepath)
                    lst_path_obrigacao.append(dirobrig)
                    
    print("")            
    log("PASTAS QUE ATENDEM AOS PARAMETROS INFORMADOS:")
    lst_obrig_mov = []
    for p in lst_path_obrigacao:
        maskobrigold = p.replace("OBRIGACAO",configuracoes.OBRIGACAO)
        #print("Mascara obrigacao_old = ", maskobrigold)
        obrig_old = glob.glob(maskobrigold)
        if (len(obrig_old) == 0):
            lst_obrig_mov.append(p)
            log("Pasta: ", p)            
            
    print("")        
    for dirobrig in lst_obrig_mov:
        #print("dirobrig = ",dirobrig)
        dirobrigold = dirobrig.replace("OBRIGACAO",configuracoes.OBRIGACAO)
        
        try:
            log("Movendo ", dirobrig, " para ", dirobrigold)
            shutil.move(dirobrig,dirobrigold)

            log("Recriando o diretorio: ", dirobrig)
            os.mkdir(dirobrig)
            
            log("Copiando M,I,D de ULTIMA_ENTREGA PARA OBRIGACAO...")
            origem  = dirobrig.replace("OBRIGACAO","ULTIMA_ENTREGA")
            destino = dirobrig
                
            maskfiles=os.path.join(origem,"*[M,I,D].[0-9][0-9][0-9]")
            #print(maskfiles)

            for file in glob.glob(maskfiles):
                nomearq = file.split(SD)[-1].split(".")[-2]
                #print("Nome do aquivo = ", nomearq)
                anoatu = int(file.split(SD)[-8])
                #print("anoatu = ", anoatu)
                
                lennomearq = len(nomearq)
                
                
                #print("lennomearq = ", lennomearq)
                #print("nomearq = ", nomearq)
                #print("anoatu = ", anoatu)
                
                if (lennomearq == 29 and anoatu > 16):
                    serienome = nomearq[-11:-8] 
                elif (lennomearq == 11 and anoatu < 17):
                    serienome = nomearq[-9:-6] 
                else:
                    serienome = "ERR"
                    log("ERRO - Arquivo com nome inválido encontrado na pasta ULTIMA_ENTREGA: ", file)
                    ret = 55
                    continue
                #print("serienome = ", serienome)
                if (serienome in ("C  ","1  ","UT ","UK ","Z01","Z02","Z03","Z04" ) or serienome.startswith("V")):
                    log("ERRO - serie ", serienome ," encontrada na pasta ULTIMA_ENTREGA")
                    ret = 55
                    continue
                
                destino = file.replace("ULTIMA_ENTREGA","OBRIGACAO")
                log("Copiando \n", file, "\n para \n", destino, " \n")
                
                shutil.copy2(file,destino)
   
        except Exception:    
            log("Erro - Não foi possível copiar o arquivo ", file , " para ", destino )
            ret = 99

    return ret
      
 
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
        ret = processar(CC_IE,CC_FILIAL,DT_MESANO_INICIO,DT_MESANO_FIM,CC_SERIE,UFI)
        if ( ret != 0) :
            log('ERRO no processamento ... Verifique. RC = ', ret)
    sys.exit(ret)



