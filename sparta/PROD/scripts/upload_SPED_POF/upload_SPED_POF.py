#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-

"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: upload_NPGF.py
  CRIACAO ..: 21/05/2020
  AUTOR ....: Welber Pena de Sousa / KYROS Consultoria
  DESCRICAO : 
                

  ANEXO ....: Demais dados e documentação na pasta documentação, arquivos :
                - Teshuva_EspecificaçãoFuncionalProcUpload_EliminarAçãoManual_v1.docx
                - Teshuva_EspecificaçãoFuncionalProcUpload_EliminarAçãoManual_v2.docx
                - Teshuva_EspecificaçãoFuncionalProcUpload_EliminarAçãoManual_v3.docx
                - Teshuva_EspecificaçãoFuncionalProcUpload_EliminarAçãoManual_v4.docx
                - Teshuva_EspecificaçãoFuncionalProcUpload_EliminarAçãoManual_v5.docx
                - Teshuva_EspecificaçãoFuncionalProcUpload_EliminarAçãoManual_v6.docx

----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 21/05/2020 - Welber Pena de Sousa - Kyros Consultoria
        - Criacao do script.
        
    * 16/05/2020 - Airton Borges da Silva Filho - Kyros Consultoria 
        - Alterações no processo de upload do SPED
            - padronização das pastas
            - aceitar parametro IE opcional
            - selecionar somente o arquivo mais novo
    * 02/02/2021 - Airton Borges da Silva Filho - Kyros Consultoria 
        - Alterações no processo de upload do GIA
            - padronização das pastas
            - aceitar parametro IE opcional
            - selecionar somente o arquivo mais novo      
        
    * 03/09/2021    
        Adequação para novo formato de script 
        SCRIPT ......: loader_sped_registro_O150.py
        AUTOR .......: Victor Santos    

    01/02/2021 - ALT001 - Welber Pena
        - Alterado o diretorio de relatorios dos arquivos GIA e suas mascaras de relatorios 
            a realizar o Upload.
        - Retirada a mascara de relatorios do SPED a carregar 
            - maska = 'Analise_SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_*.xlsx'
    - 24/02/2022 - Eduardo da Silva Ferreira - Kyros Tecnologia
            - [PTITES-1635] Padrão de diretórios do SPARTA

----------------------------------------------------------------------------------------------
"""

import os
import sys

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)
import configuracoes

import fnmatch

import time
import base64
from urllib.parse import urlencode, quote_plus
import http
import requests
import urllib.parse
import uuid
import glob
from pathlib import Path
import datetime

from funcs_upload import *

variaveis = {}

import comum
import sql
import layout
import util

log.gerar_log_em_arquivo = True
comum.carregaConfiguracoes(configuracoes)
name_script = os.path.basename(__file__).split('.')[0]

def verificaArquivo(x, pos, letra) :
    if len(x) > pos :
        if x[pos] == letra :
            return True
    return False

def ies_existentes(mascara,diretorio):
    global ret
   
    qdade = 0
    ies = []
    directory = Path(diretorio)
    files = directory.glob(mascara)
    sorted_files = sorted(files, reverse=False)

    if sorted_files:
        log("# Arquivos encontrados: ")
        for f in sorted_files:
            qdade = qdade + 1
            pos = 4
            if (mascara != 'RELATORIOS.zip'):
                if( str(f).split('.')[-1] == 'xlsx'):
                    pos = 5
                ie = str(f).split("_")[pos]
                log("#   ",qdade, " => ", f, " IE = ", ie)
                try:
                    ies.index(str(f).split("_")[pos])
                except:
                    ies.append(str(f).split("_")[pos])
                    continue
            else:
                log("#   ",qdade, " => ", f)
    else: 
        log('#### ERRO:    Arquivo %s não está na pasta %s'%(mascara,diretorio))
        ret=99
        return("")
    log("-"*100)
    return(ies)

def ies_GIA_existentes(mascara,diretorio):
    global ret
   
    qdade = 0
    ies = []
    directory = Path(diretorio)
    log (directory)
    files = directory.glob(mascara)
    sorted_files = sorted(files, reverse=False)
    log(sorted_files)
    if sorted_files:
        log("# Arquivos encontrados: ")
        for f in sorted_files:
            qdade = qdade + 1
            pos = 2
            if (mascara != 'RELATORIOS.zip'):
                if( str(f).split('.')[-1] == 'xlsx'):
                    pos = 4
                ie = str(f).split("_")[pos]
                log("#   ",qdade, " => ", f, " IE = ", ie)
                try:
                    ies.index(str(f).split("_")[pos])
                except:
                    ies.append(str(f).split("_")[pos])
                    continue
            else:
                log("#   ",qdade, " => ", f)
    else: 
        log('Aqui 2')
        log('#### ERRO:    Arquivo %s não está na pasta %s'%(mascara,diretorio))
        ret=99
        return("")
    log("-"*100)
    return(ies)

def nome_arquivo(mascara,diretorio):
    qdade = 0
    nomearq = "" 
    directory = Path(diretorio)
    files = directory.glob(mascara)
    
    log("  = ", )
    log("mascara  = ",mascara )
    log("diretorio  = ",diretorio )
    log("directory  = ",directory )
    log("files  = ",files )
    log("  = ", )
    log("  = ", )

    sorted_files = sorted(files, reverse=False)
    if sorted_files:
        for f in sorted_files:
            qdade = qdade + 1
            nomearq = os.path.basename(f)
    else: 
        log("-"*100)
        log('Aqui 3')
        log('#### ERRO:    Arquivo %s não está na pasta %s'%(mascara,diretorio))
        log("-"*100)
    return(nomearq)

def carregaDicionarioEmpresa() :
    arq_empresas = 'empresasFiliais.cfg'
    path_arq_empresas = os.path.join( '.', arq_empresas )
    empresas = {}
    
    if os.path.isfile( path_arq_empresas ) :
        log('Carregando dados do arquivo de configuracao .:', arq_empresas)
        log('- Path .:', path_arq_empresas)
        fd = open(path_arq_empresas, 'r')
        linhas = fd.readlines()

        for item in linhas[1:] :
            if not item.startswith('#') :
                if item.__contains__(';') :
                    linha = item.replace('\n','').split(';')
                    if len(linha) >= 2 :
                        empresas[linha[0].strip()] = empresas.get(linha[0].strip(), {})
                        empresas[linha[0].strip()][linha[1].strip()] = [linha[2].strip(), linha[3].strip()]
                        log('- Para tipo %s e estado %s : Empresa %s - Filial %s'%(linha[0].strip(), linha[1].strip(), linha[2].strip(), linha[3].strip()) )
        fd.close()
    
    return empresas

def realizaUpload() :
    log('Iniciando processo de UPLOAD ...')
    empresas = carregaDicionarioEmpresa()
    ### de acordo com o tipo de arquivos a tratar define os diretorios de trabalho .
    diretorios_trabalho = {}
    diretorios_trabalhox = {}
    total_arqs_transferidos = 0
    
    if configuracoes.tipo_arquivo == 'conv115' :
        id_obrigacao = 1
        if(configuracoes.envia_enxertado == 'S'):        
            #### Diretorio de Obrigação
            diretorios_trabalho['OBRIGACAO'] = {}
            diretorios_trabalho['OBRIGACAO']['diretorio'] = os.path.join( configuracoes.caminho_LEVCV115,configuracoes.uf, str(configuracoes.ano)[-2:], configuracoes.mes, 'TBRA', configuracoes.filial, 'SERIE', configuracoes.id_serie, 'OBRIGACAO' )
            diretorios_trabalho['OBRIGACAO']['mascara'] = ['*']
            diretorios_trabalho['OBRIGACAO']['tipo'] = 'regerados'
        
        if(configuracoes.envia_protocolado == 'S'):
            #### Diretorio de Obrigação
            diretorios_trabalho['PROTOCOLADO'] = {}
            diretorios_trabalho['PROTOCOLADO']['diretorio'] = os.path.join( configuracoes.caminho_LEVCV115,configuracoes.uf, str(configuracoes.ano)[-2:], configuracoes.mes, 'TBRA', configuracoes.filial, 'SERIE', configuracoes.id_serie, 'OBRIGACAO' )
            diretorios_trabalho['PROTOCOLADO']['mascara'] = ['*']
            diretorios_trabalho['PROTOCOLADO']['tipo'] = 'originais'

        if(configuracoes.envia_relatorio == 'S'):
            #### Diretorio de INSUMOS
            diretorios_trabalho['INSUMOS'] = {}
            diretorios_trabalho['INSUMOS']['diretorio'] = os.path.join( configuracoes.caminho_LEVCV115,configuracoes.uf, str(configuracoes.ano)[-2:], configuracoes.mes, 'TBRA', configuracoes.filial, 'SERIE', configuracoes.id_serie, 'INSUMOS' )
            diretorios_trabalho['INSUMOS']['mascara'] = ['Conciliacao_Serie_*OK.xlsx']
            diretorios_trabalho['INSUMOS']['tipo'] = 'analises'

            #### Diretorio de PVA
            diretorios_trabalho['PVA'] = {}
            diretorios_trabalho['PVA']['diretorio'] = os.path.join( configuracoes.caminho_LEVCV115,configuracoes.uf, str(configuracoes.ano)[-2:], configuracoes.mes, 'TBRA', configuracoes.filial, 'SERIE', configuracoes.id_serie, 'PVA' )
            diretorios_trabalho['PVA']['mascara'] = ['Log*.txt', 'Erros*.txt']
            diretorios_trabalho['PVA']['tipo'] = 'analises'

    elif configuracoes.tipo_arquivo.upper() == 'SPED' :
        arquivo_compactado = ""
        id_obrigacao = 2
        configuracoes.uf = configuracoes.uf.upper()
        maskp = 'SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_PROT'+'*.txt'
        maskr = 'SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_REG'+'*.txt'
        maske = 'SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_ENX'+'*.txt'
        ### ALT001 - Inicio
        ###     - Retirada a mascara de relatorios a enviar abaixo 
        ##maska = 'Analise_SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_*.xlsx'
        ### ALT001 - Fim
        maskac = 'RELATORIOS.zip'
        
        #### Diretorio de SPED - Protocolados
        if(configuracoes.envia_protocolado == 'S'):
            diretorios_trabalho['SPED'] = {}
            # [PTITES-1635] # diretorios_trabalho['SPED']['diretorio'] = os.path.join( configuracoes.caminho_sped_fiscal, 'PROTOCOLADOS', configuracoes.uf, configuracoes.ano, configuracoes.mes )
            diretorios_trabalho['SPED']['diretorio'] = os.path.join( os.path.dirname(configuracoes.dir_entrada), 'SPED_FISCAL', 'PROTOCOLADOS', configuracoes.uf, configuracoes.ano, configuracoes.mes )
            diretorios_trabalho['SPED']['mascara'] = [maskp]
            diretorios_trabalho['SPED']['tipo'] = 'originais'
        
        #### Diretorio de SPED - Enxertados
        if(configuracoes.envia_enxertado == 'S'):        
            diretorios_trabalho['ENXERTADOS'] = {}
            # [PTITES-1635] # diretorios_trabalho['ENXERTADOS']['diretorio'] = os.path.join( configuracoes.caminho_sped_fiscal, 'ENXERTADOS', configuracoes.uf,  configuracoes.ano, configuracoes.mes )
            diretorios_trabalho['ENXERTADOS']['diretorio'] = os.path.join( os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'ENXERTADOS', configuracoes.uf,  configuracoes.ano, configuracoes.mes )
            diretorios_trabalho['ENXERTADOS']['mascara'] = [maske]
            diretorios_trabalho['ENXERTADOS']['tipo'] = 'regerados'
       
        #### Diretorio de SPED - Relatórios
        if(configuracoes.envia_relatorio == 'S'):
            diretorios_trabalho['RELATORIOS'] = {}
            ### ALT001 - Inicio - Retira a mascara abaixo .
                ### ALT001 - Alterado o diretorio de arquivos do Relatorio_Conciliacao_*.xlsx
            # diretorios_trabalho['RELATORIOS']['diretorio'] = os.path.join( configuracoes.caminho_relatorios, 'SPED_FISCAL', configuracoes.uf,  configuracoes.ano, configuracoes.mes )
            diretorios_trabalho['RELATORIOS']['diretorio'] = os.path.join( configuracoes.dir_geracao_arquivos.split('upload')[0], 'Insumos', 'SPED_FISCAL', configuracoes.uf,  configuracoes.ano, configuracoes.mes )
            
            diretorios_trabalho['RELATORIOS']['mascara'] = [
                                                            ### ALT001 - Retira a mascara abaixo .
                                                            #maska, 
                                                            ### ALT001 - Acrescentada a mascara abaixo
                                                            'Detalhamento_Obrigações_Fiscais_*.xlsm',
            ### ALT001 - Fim
                                                            'Relatorio_Conciliacao_*.xlsx'
                                                        ]
            diretorios_trabalho['RELATORIOS']['tipo'] = 'analises'
            diretorios_trabalhox['APOIO'] = {}
            diretorios_trabalhox['APOIO']['diretorio'] = os.path.join( diretorios_trabalho['RELATORIOS']['diretorio'], 'Apoio' )
            diretorios_trabalhox['APOIO']['mascara'] = 'RELATORIOS.zip'
            
            arquivo_compactado = os.path.join((diretorios_trabalho['RELATORIOS']['diretorio'] ) , 'RELATORIOS.zip') + " " 
            
            os.system('rm ' + arquivo_compactado)
            
            compacta = "zip -j -q -n zip " + arquivo_compactado  + " "
            for masc in diretorios_trabalho['RELATORIOS']['mascara'] :
                arq_rel = nome_arquivo(masc, diretorios_trabalho['RELATORIOS']['diretorio'])
                if arq_rel:
                    compacta = compacta + os.path.join((diretorios_trabalho['RELATORIOS']['diretorio']), arq_rel) + " "

            compacta = compacta + os.path.join((diretorios_trabalhox['APOIO']['diretorio']), '*.*')

            os.system(compacta)
            
            diretorios_trabalho['RELATORIOS']['mascara'] = [maskac]


        configuracoes.empresa = empresas.get(configuracoes.tipo_arquivo,{}).get(configuracoes.uf,[False,False])[0]
        configuracoes.filial = empresas.get(configuracoes.tipo_arquivo,{}).get(configuracoes.uf,[False,False])[1]

        # variaveis['empresa'] = empresas.get(variaveis['tipo_arquivo'],{}).get(variaveis['uf'],[False,False])[0]
        # variaveis['filial'] = empresas.get(variaveis['tipo_arquivo'],{}).get(variaveis['uf'],[False,False])[1]
        
        log('configuracoes.empresa',configuracoes.empresa)

        if not configuracoes.empresa or not configuracoes.filial :
            log('ERRO - SPED Falta configuracao de empresa e filial no arquivo empresasFiliais.cfg')
            return False
        # 'TBRA' e '0001'
     
    elif configuracoes.tipo_arquivo.upper() == 'GIA' :
        arquivo_compactado = ""
        id_obrigacao = 3
        configuracoes.uf = configuracoes.uf.upper()
        ano = configuracoes.ano
        # ano = int(ano) + 2000
        # ano = str(ano)
        maskp = 'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.ie+'_PROT'+'*.txt'
        maskr = 'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.ie+'_REG'+'*.txt'
        maske = 'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.ie+'_ENX'+'*.txt'
        maskac = 'RELATORIOS.zip'
        maska = 'relatorio_gia_totais_por_cfop_'+str(configuracoes.mes)+str(ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_*.xlsx'

        #### Diretorio de SPED - Protocolados
        if(configuracoes.envia_protocolado == 'S'):
            diretorios_trabalho['GIA'] = {}
            # [PTITES-1635] # diretorios_trabalho['GIA']['diretorio'] = os.path.join(configuracoes.caminho_GIA+configuracoes.uf,'PROTOCOLADOS',ano,configuracoes.mes)
            diretorios_trabalho['GIA']['diretorio'] = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'GIA', 'PROTOCOLADOS', configuracoes.uf, ano,configuracoes.mes)
            diretorios_trabalho['GIA']['mascara'] = [maskp]
            diretorios_trabalho['GIA']['tipo'] = 'originais'

        #### Diretorio de SPED - Enxertados
        if(configuracoes.envia_enxertado == 'S'):        
            diretorios_trabalho['ENXERTADOS'] = {}
            # [PTITES-1635] # diretorios_trabalho['ENXERTADOS']['diretorio'] = os.path.join(configuracoes.caminho_GIA+configuracoes.uf,'ENXERTADOS',ano,configuracoes.mes)
            diretorios_trabalho['ENXERTADOS']['diretorio'] = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'GIA','ENXERTADOS', configuracoes.uf,ano,configuracoes.mes)
            diretorios_trabalho['ENXERTADOS']['mascara'] = [maske]
            diretorios_trabalho['ENXERTADOS']['tipo'] = 'regerados'

        #### Diretorio de SPED - Relatórios
        if(configuracoes.envia_relatorio == 'S'):
            diretorios_trabalho['RELATORIOS'] = {}
            ### ALT001 - Inicio
                ### ALT001 - Alterado o diretorio de entrada dos relatorios .
            # diretorios_trabalho['RELATORIOS']['diretorio'] = os.path.join(configuracoes.caminho_relatorios,'GIA',configuracoes.uf,ano,configuracoes.mes)
            diretorios_trabalho['RELATORIOS']['diretorio'] = os.path.join(configuracoes.dir_geracao_arquivos.split('upload')[0] , 'Insumos', 'GIA', configuracoes.uf, configuracoes.ano, configuracoes.mes )
            diretorios_trabalho['RELATORIOS']['mascara'] = [ 
                                                            
                                                            ### ALT001 - Retiradas as 3 mascaras abaixo
                                                            #maska, 
                                                            #'Analise_SPED_Convenio115_GIA_%s_%s_*'%( str(configuracoes.mes)+str(ano), configuracoes.uf ), 
                                                            #'Insumo_CFOP_SPED_GIA_%s_%s_*'%( str(configuracoes.mes)+str(ano), configuracoes.uf ) ,
                                                            '*.pdf'
            ### ALT001 - Fim
                                                        ]
            diretorios_trabalho['RELATORIOS']['tipo'] = 'analises'

        configuracoes.empresa = empresas.get(configuracoes.tipo_arquivo,{}).get(configuracoes.uf,[False,False])[0]
        configuracoes.filial = empresas.get(configuracoes.tipo_arquivo,{}).get(configuracoes.uf,[False,False])[1]
        
        if not configuracoes.empresa or not configuracoes.filial :
            log('ERRO - GIA Falta configuracao de empresa e filial no arquivo empresasFiliais.cfg')
            return False
        # 'TBRA' e '0001'

    if configuracoes.tipo_arquivo.upper() == 'SEF' :
        id_obrigacao = 3
         #### Diretorio de GIA - Protocolados
        diretorios_trabalho['SEF'] = {}
        diretorios_trabalho['SEF']['diretorio'] = os.path.join( '/arquivos', 'SEF2', configuracoes.uf, 'PROTOCOLADOS', configuracoes.sub_dir )
        diretorios_trabalho['SEF']['mascara'] = ['*']
        diretorios_trabalho['SEF']['tipo'] = 'originais'

        #### Diretorio de GIA - Enxertados
        diretorios_trabalho['ENXERTADOS'] = {}
        diretorios_trabalho['ENXERTADOS']['diretorio'] = os.path.join( '/arquivos', 'SEF2', configuracoes.uf, 'ENXERTADOS', configuracoes.sub_dir )
        diretorios_trabalho['ENXERTADOS']['mascara'] = ['ENXERTADO_*']
        diretorios_trabalho['ENXERTADOS']['tipo'] = 'regerados'
        
        configuracoes.empresa = configuracoes.tipo_arquivo,{}.get(configuracoes.uf,[False,False])[0]
        configuracoes.filial = configuracoes.tipo_arquivo,{}.get(configuracoes.uf,[False,False])[1]
        
        if not configuracoes.empresa or not configuracoes.filial :
            log('ERRO - SEF - Falta configuracao de empresa e filial no arquivo empresasFiliais.cfg')
            return False
    
    lst_protocolados = []
    lst_recibos = []
    lst_regerados = []
    lst_analise = []
    lst_keys_diretorios_trabalho = [ x for x in diretorios_trabalho.keys() ]
    
    for nome_dir in lst_keys_diretorios_trabalho :
        if (nome_dir == 'APOIO'):
            continue
        log('Realizando upload dos arquivos %s de %s.'%( configuracoes.tipo_arquivo, nome_dir ))
        dir_trabalho = diretorios_trabalho[nome_dir]['diretorio']
        mascaras = diretorios_trabalho[nome_dir]['mascara']
        log('='*100)
        log('Diretorio       : %s'%( dir_trabalho ))
        log('Nome do Arquivo : %s'%( mascaras ))
        
        if os.path.isdir( dir_trabalho ) :
            log('- Mascaras a enviar .: %s'%( ' '.join( x for x in mascaras ) )) 
      
            listaarquivos = os.listdir(dir_trabalho)

            if configuracoes.tipo_arquivo == 'SPED' :  
                listaarquivos = []
                listadeies = ies_existentes(mascaras, dir_trabalho) 
    
                for iee in listadeies:
                     
                    mascarass = ""
 
                    if (mascaras == 'SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_PROT'+'*.txt'):
                        mascarass = 'SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+iee+'_PROT'+'*.txt'   

    
                    if (mascaras == 'SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_REG'+'*.txt'):
                        mascarass = 'SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+iee+'_REG'+'*.txt'

                        
                    if (mascaras == 'SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_ENX'+'*.txt'):
                        mascarass = 'SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+iee+'_ENX'+'*.txt'

                        
                    if (mascaras == 'Analise_SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_*.xlsx'):
                        mascarass = 'Analise_SPED_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+iee+'*.xlsx'

                    listaarquivos.append(nome_arquivo(mascarass,dir_trabalho))

                if (mascaras == 'RELATORIOS.zip'):
                    mascarass = 'RELATORIOS.zip'
                    listaarquivos.append(nome_arquivo(mascarass,dir_trabalho))


            elif configuracoes.tipo_arquivo == 'GIA' :  
                listaarquivos = []
                listadeies = ies_GIA_existentes(mascaras, dir_trabalho) 
    
                for iee in listadeies:
                     
                    mascarass = ""
 
                    if (mascaras == 'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.ie+'_PROT'+'*.txt'):
                        mascarass = 'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+iee+'_PROT'+'*.txt'   

    
                    if (mascaras == 'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.ie+'_REG'+'*.txt'):
                        mascarass = 'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+iee+'_REG'+'*.txt'

                        
                    if (mascaras == 'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.ie+'_ENX'+'*.txt'):
                        mascarass = 'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+iee+'_ENX'+'*.txt'

                        
                    if (mascaras == 'Comando_SO_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_*.xlsx'):
                        mascarass = 'Comando_SO_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+iee+'*.xlsx'



                    log(" = ", 'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.ie+'_PROT'+'*.txt')
                    log(" = ",'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.ie+'_REG'+'*.txt' )
                    log(" = ",'GIA'+configuracoes.uf+'_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.ie+'_ENX'+'*.txt' )
                    log(" = ",'Comando_SO_'+str(configuracoes.mes)+str(configuracoes.ano)+'_'+configuracoes.uf+'_'+configuracoes.ie+'_*.xlsx' )
                    log(" = ", )                    
                    log("mascaras = ",mascaras )
                    log("mascarass = ",mascarass )
                    log("iee = ",iee )
                    log("dir_trabalho = ",dir_trabalho )
                    log("listadeies = ",listadeies )
                    log(" = ", )
                    log(" = ", )
                    log(" = ", )

                    listaarquivos.append(nome_arquivo(mascarass,dir_trabalho))

                if (mascaras == 'RELATORIOS.zip'):
                    mascarass = 'RELATORIOS.zip'
                    listaarquivos.append(nome_arquivo(mascarass,dir_trabalho))
            
            for arq in listaarquivos :    
                
                if (arquivoAEnviar(arq, mascaras) or mascaras == 'RELATORIOS.zip'):
                    total_arqs_transferidos += 1
                    log('-'*80)
                    log('Transferindo o arquivo ..: %s'%(arq))
                   
                    autenticacao = authenticate()
                    if not autenticacao :
                        return False
                    log('- Usuario : %s - ID : %s'%( autenticacao.get('displayName'), autenticacao.get('id') ))
                    dic_args = enviarArquivo(autenticacao, dir_trabalho, arq)
                    if not dic_args :
                        return False
                    
                    status = os.stat(os.path.join(dir_trabalho,arq))
                    
                    dt_ts = datetime.datetime.fromtimestamp(status.st_mtime, tz=datetime.timezone.utc)
                    
                    FileRepresentation = {}
                    FileRepresentation["originalName"] = arq
                    FileRepresentation["internalName"] = dic_args.get('resumableIdentifier', '') #### "{guid-v4}" // Usar o mesmo valor utilizado no parâmetro resumableIdentifier durante o upload
                    FileRepresentation["strongHash"] = dic_args["hash-md5"]
                    FileRepresentation["originalSize"] = dic_args['resumableTotalSize']  #### "{tamanho-total-do-arquivo-em-bytes}"
                    FileRepresentation["lastModifiedDate"] = dt_ts.isoformat() ### // Informar a data no formato UTC string (conforme a RFC-11

                    if diretorios_trabalho[nome_dir]['tipo'] == 'originais' :
                        lst_protocolados.append(FileRepresentation)
                    elif diretorios_trabalho[nome_dir]['tipo'] == 'analises' :
                        lst_analise.append(FileRepresentation)
                    elif diretorios_trabalho[nome_dir]['tipo'] == 'regerados' :
                        lst_regerados.append(FileRepresentation)
                    elif diretorios_trabalho[nome_dir]['tipo'] == 'recibos' :
                        lst_recibos.append(FileRepresentation)
        
        else :
            log('Erro - Diretorio de arquivos não existe, favor validar !')

    if not log.ret :
        #### Depois de realizar o Upload dos arquivos (tópico anterior), é necessário registrar o envio no Teshuvá, 
        ##   para isso, realize um POST /api/teshuva/controle
        if(total_arqs_transferidos > 0) : 
            if not registraUpload(autenticacao, id_obrigacao, lst_protocolados, lst_recibos, lst_regerados, lst_analise ) :
                log('ERRO ao registrar uploads no portal.')
                return False
        else:
            log('#### ERRO #### Nenhum arquivo transferido.')
            log.ret = 10
            return False

    log('  R E S U M O  '.center(50,'*')) 
    log('-'* 100)
    log('-'* 100)
    log('* PROTOCOLADOS TRANSFERIDOS ....  \n', lst_protocolados )
    log('-'* 100)
    log('* REGERADOS TRANSFERIDOS .......  \n', lst_regerados )
    log('-'* 100)
    log('* ANALISES TRANSFERIDOS ......  \n', lst_analise )
    log('-'* 100)
    log('* TOTAL DE ARQUIVOS TRANSFERIDOS = < %s > arquivos.'%( total_arqs_transferidos ))
    log('-'* 100)
    log('-'* 100)
    log('*'*50)

    return True


def arquivoAEnviar(arq, mascaras) :
    valido = False
    for mascara in mascaras :
        if fnmatch.fnmatch(arq,mascara) :
            valido = True
    return valido

def processar() :
    global variaveis
    #####
    ## Identifica o(s) parametro(s) do script :
    ##     1 - < id_serie > - obrigatorio
    ##     2 - < Tipo_arquivo > 
    try :

        if len(sys.argv) > 2 :
           
            configuracoes.tipo_arquivo = sys.argv[1].lower()

            if configuracoes.tipo_arquivo not in ['conv115', 'sped', 'gia', 'sef'] :
                log.ret = 99
            else :

                if configuracoes.tipo_arquivo == 'conv115' :
                    #### Parametros para o conv115
                    configuracoes.id_serie = sys.argv[7]
                    configuracoes.envia_protocolado = ('N' if(sys.argv[4].upper() != 'S') else 'S') 
                    configuracoes.envia_enxertado = ('N' if(sys.argv[5].upper() != 'S') else 'S') 
                    configuracoes.envia_relatorio = ('N' if(sys.argv[6].upper() != 'S') else 'S') 

                elif configuracoes.tipo_arquivo == 'sped' :
                    if (len(sys.argv) < 7):
                        log.ret = 99
                    elif(sys.argv[4].upper() not in ['N', 'S'] or sys.argv[5].upper() not in ['N', 'S'] or sys.argv[6].upper() not in ['N', 'S']):
                        log.ret = 99
                    else:
                        #### Parametros para o sped
                        # configuracoes.tipo_arquivo = sys.argv[2].lower()
                        configuracoes.uf = sys.argv[2]
                        configuracoes.sub_dir = sys.argv[3]
                        configuracoes.mes, configuracoes.ANO = sys.argv[3][:2],sys.argv[3][2:] 
                        # configuracoes.ano = configuracoes.ANO[-2:]
                        configuracoes.ano = configuracoes.ANO
                        configuracoes.envia_protocolado = ('N' if(sys.argv[4].upper() != 'S') else 'S') 
                        configuracoes.envia_enxertado = ('N' if(sys.argv[5].upper() != 'S') else 'S') 
                        configuracoes.envia_relatorio = ('N' if(sys.argv[6].upper() != 'S') else 'S') 
                        
                        if (configuracoes.envia_protocolado == configuracoes.envia_enxertado == configuracoes.envia_relatorio == "N"):
                            log.ret = 99    
                            log('Erro encontrado : NENUM TIPO DE ARQUIVO FOI MARCADO PARA UPLOAD.')
                        
                        configuracoes.ie = ('*' if len(sys.argv) < 8 else sys.argv[7])
                        if(configuracoes.ie == ""):
                            configuracoes.ie = "*"
                        
                        if (configuracoes.ie != '*'):
                            if (int(configuracoes.ie) < 100):
                                log.ret = 99    
                                log('Erro encontrado : IE inválido.')
                    
                elif configuracoes.tipo_arquivo == 'sef' :
                    #### Parametros para o SEFII
                    configuracoes.uf = sys.argv[2]
                    configuracoes.sub_dir = sys.argv[3]
                    configuracoes.mes, configuracoes.ANO = sys.argv[3][:2],sys.argv[3][2:] 
                    # configuracoes.ano = configuracoes.ANO[-2:]
                    configuracoes.ano = configuracoes.ANO

                else :
                    if (len(sys.argv) < 7):
                        log.ret = 99
                    elif(sys.argv[4].upper() not in ['N', 'S'] or sys.argv[5].upper() not in ['N', 'S'] or sys.argv[6].upper() not in ['N', 'S']):
                        log.ret = 99
                    else:
                        #### Parametros para o sped
                        configuracoes.uf = sys.argv[2]
                        configuracoes.sub_dir = sys.argv[3]
                        configuracoes.mes,  configuracoes.ANO = sys.argv[3][:2],sys.argv[3][2:] 
                        configuracoes.ano = configuracoes.ANO
                        # configuracoes.ano = configuracoes.ANO[-2:]
                        configuracoes.envia_protocolado = ('N' if(sys.argv[4].upper() != 'S') else 'S') 
                        configuracoes.envia_enxertado = ('N' if(sys.argv[5].upper() != 'S') else 'S') 
                        configuracoes.envia_relatorio = ('N' if(sys.argv[6].upper() != 'S') else 'S') 
                        
                        if (configuracoes.envia_protocolado == configuracoes.envia_enxertado == configuracoes.envia_relatorio == "N"):
                            log.ret = 99    
                            log('Erro encontrado : NENUM TIPO DE ARQUIVO FOI MARCADO PARA UPLOAD.')
                        
                        configuracoes.ie = ('*' if len(sys.argv) < 8 else sys.argv[7])
                        if(configuracoes.ie == ""):
                            configuracoes.ie = "*"
                        
                        if (configuracoes.ie != '*'):
                            if (int(configuracoes.ie) < 100):
                                log.ret = 99    
                                log('Erro encontrado : IE inválido.')                    
        else :
            log.ret = 99
    except Exception as e :
        log('Erro encontrado :', e)
        log.ret = 99
    
    if log.ret == 99 :
        log('ERRO - Erro nos parametros do script.')
        log('Exemplo :')
        log('     ./%s.py <Tipo_aquivo> <outros parametros>'%(name_script))
        log('     ')
        log('     <Tipo_arquivo> pode ter os seguintes valores :')
        log('           conv115, sped ou gia')
        log('     os demais parametros sao de acordo com o <Tipo_arquivo>')
        log('          - Caso tipo_arquivo = conv115 os parametros sao :')
        log('               1 - <uf>        // * nao utilizado passar "" ')
        log('               2 - <mesano>    // * nao utilizado passar "" ')
        log('               3 - <S> ou <N>  // * obrigatorio - <S> para fazer ou <N> para não fazer o upload do PROTOCOLADO')
        log('               4 - <S> ou <N>  // * obrigatorio - <S> para fazer ou <N> para não fazer o upload do ENXERTADO')
        log('               5 - <S> ou <N>  // * obrigatorio - <S> para fazer ou <N> para não fazer o upload do RELATORIO E arquivos de Apoio')
        log('               6 - <id_serie>  // * obrigatorio - no formato 99999999')  
        log('            Exemplo : ./upload_NPGF.py conv115 "" "" N N S 18000394')
        log('')
        log('          - Caso tipo_arquivo = sped os parametros sao :')
        log('               1 - <uf>        // * obrigatorio ')
        log('               2 - <mesano>    // * obrigatorio - no formato mmyyyy')
        log('               3 - <S> ou <N>  // * obrigatorio - <S> para fazer ou <N> para não fazer o upload do PROTOCOLADO')
        log('               4 - <S> ou <N>  // * obrigatorio - <S> para fazer ou <N> para não fazer o upload do ENXERTADO')
        log('               5 - <S> ou <N>  // * obrigatorio - <S> para fazer ou <N> para não fazer o upload do RELATORIO E arquivos de Apoio')
        log('               6 - <IE>        // * opcional - no formato 999999999999')       
        log('            Exemplo 1 : ./upload_NPGF.py sped SP 012018 S S S 108383949112')
        log('            Exemplo 2 : ./upload_NPGF.py sped SP 012018 N N S')
        log('            Exemplo 3 : ./upload_NPGF.py sped SP 012018 S N N 108383949112')
        log('')
        log('          - Caso tipo_arquivo = gia os parametros sao :')
        log('               1 - <uf>        // * obrigatorio ')
        log('               2 - <mesano>    // * obrigatorio - no formato mmyyyy')
        log('               3 - <S> ou <N>  // * obrigatorio - <S> para fazer ou <N> para não fazer o upload do PROTOCOLADO')
        log('               4 - <S> ou <N>  // * obrigatorio - <S> para fazer ou <N> para não fazer o upload do ENXERTADO')
        log('               5 - <S> ou <N>  // * obrigatorio - <S> para fazer ou <N> para não fazer o upload do RELATORIO E arquivos de Apoio')
        log('               6 - <IE>        // * opcional - no formato 999999999999')       
        log('            Exemplo 1 : ./upload_NPGF.py gia SP 012018 S S S 108383949112')
        log('            Exemplo 2 : ./upload_NPGF.py gia SP 012018 N N S')
        log('            Exemplo 3 : ./upload_NPGF.py gia SP 012018 S N N 108383949112')
        log('')

    if not log.ret :
        conexao = sql.geraCnxBD(configuracoes)
        # conexao.autocommit = False
        # cursor = conexao.cursor()
        # variaveis['conexao'] = conexao
        # variaveis['cursor'] = cursor
        # log('- Conectado ao banco de dados!')

    if not log.ret :
        if configuracoes.tipo_arquivo == 'conv115' :
            log('Id da serie a processar ...: %s'%( configuracoes.id_serie ))
        elif configuracoes.tipo_arquivo == 'sped' :
            log('UF ......................: %s'%( configuracoes.uf ))
            log('Ano/Mes de referencia ...: %s'%( configuracoes.sub_dir ))
            log('IE ......................: %s'%( 'Todas' if configuracoes.ie == '*' else configuracoes.ie ))
        try :
            if configuracoes.tipo_arquivo == 'conv115' and not log.ret :

                variaveis = comum.buscaDadosSerie(configuracoes.id_serie)
                for var in variaveis:
                    setattr(configuracoes, var, variaveis[var])
                # if not configuracoes.variaveis :
                    # log.ret = 3
            if not log.ret :
                if not realizaUpload() :
                    log.ret = 5
        except Exception as e :
            log('Erro - Exception :', e)
            log.ret = 4
    
        # log('Finalizando conexões com o banco de dados.')
        # cursor.close()
        # conexao.close()

    return log.ret


if __name__ == "__main__":
    ret = processar()
    sys.exit(ret)





