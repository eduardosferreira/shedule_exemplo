#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
MODULO ...: TESHUVA
SCRIPT ...: retorna_dados_banco.py
CRIACAO ..: 11/01/2022
AUTOR ....: Victor Santos Cardoso - KYROS TECNOLOGIA
DESCRICAO.: 
----------------------------------------------------------------------------------------------
  HISTORICO : 
    - 22/02/2022 - Eduardo da Silva Ferreira - Kyros Tecnologia
    - [PTITES-1634] Padrão de diretórios do SPARTA
----------------------------------------------------------------------------------------------
"""
import os
import sys

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)

import configuracoes
import comum
import sql
from openpyxl import Workbook, load_workbook
import openpyxl
import datetime
import shutil
from pathlib import Path
from openpyxl.utils import get_column_letter
import fnmatch

log.gerar_log_em_arquivo = True
comum.carregaConfiguracoes(configuracoes)

def identifica_ies(uf, mes_ano, tipo):
    v_diretorio = ""
    if tipo == 'Atual Ti':  
        v_diretorio_tipo = 'ENXERTADOS'
        # - [PTITES-1634] 
        v_diretorio = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'GIA','ENXERTADOS', uf, mes_ano[-4:],mes_ano[:2])
    else:
        v_diretorio_tipo = 'PROTOCOLADOS'
        # - [PTITES-1634] 
        v_diretorio = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'GIA','PROTOCOLADOS', uf, mes_ano[-4:],mes_ano[:2])
        
    v_ies = []
    # - [PTITES-1634] v_diretorio = configuracoes.diretorio_gia
    # - [PTITES-1634] v_diretorio = v_diretorio.replace('<<UF>>', uf)
    # - [PTITES-1634] v_diretorio = v_diretorio.replace('<<AAAA>>', mes_ano[-4:])
    # - [PTITES-1634] v_diretorio = v_diretorio.replace('<<MM>>', mes_ano[:2])
    # - [PTITES-1634] v_diretorio = v_diretorio.replace('<<TIPO>>', v_diretorio_tipo)

    v_mascara_arq_GIA = configuracoes.arq_gia
    v_mascara_arq_GIA = v_mascara_arq_GIA.replace('<<UF>>', uf)
    v_mascara_arq_GIA = v_mascara_arq_GIA.replace('<<MESANO>>', mes_ano)
    v_mascara_arq_GIA = v_mascara_arq_GIA.replace('<<IE>>', '*')
    v_mascara_arq_GIA = v_mascara_arq_GIA.replace('<<TIPO_ABREVIADO>>', '*')
    v_mascara_arq_GIA = v_mascara_arq_GIA.replace('<<NNN>>', '*')

    l_diretorio = Path(v_diretorio)
    l_arq = l_diretorio.glob(v_mascara_arq_GIA)
    l_procura_arquivos = sorted(l_arq, reverse=False)
    if l_procura_arquivos:        
        for arquivo in l_procura_arquivos:
            ie = os.path.basename(str(arquivo)).split('_')[2]
            if ie not in v_ies:
                v_ies.append(ie)
    return v_ies

def ultimo_Arquivo_Diretorio(p_arq_mascara, p_diretorio):

    v_arquivo = ""
    l_diretorio = Path(p_diretorio)
    l_arq = l_diretorio.glob(p_arq_mascara)
    l_procura_arquivos = sorted(l_arq, reverse=False)
    
    if l_procura_arquivos:        
        for item in l_procura_arquivos:    
            arquivo = os.path.basename(str(item))
            if fnmatch.fnmatch(arquivo,p_arq_mascara) :
                v_arquivo = arquivo
    return v_arquivo

def tipoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r', encoding='iso-8859-1')
        fd.readline()
        fd.close()
    except :
        return 'utf-8'
    return 'iso-8859-1'

def converter(uf, mes_ano, tipo):

    dados_retorno = []
    if tipo == 'Atual Ti':  
        v_diretorio_tipo = 'ENXERTADOS'
        v_mascara_tipo = 'ENX*'
        # - [PTITES-1634] 
        v_diretorio = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'GIA','ENXERTADOS', uf,mes_ano[-4:],mes_ano[:2])
    else:
        v_diretorio_tipo = 'PROTOCOLADOS'
        v_mascara_tipo = 'PROT'
        # - [PTITES-1634] 
        v_diretorio = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'GIA','PROTOCOLADOS', uf,mes_ano[-4:],mes_ano[:2])

    v_lst_ies = identifica_ies( uf, mes_ano, tipo )
    # - [PTITES-1634] -- v_diretorio = configuracoes.diretorio_gia
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<UF>>', uf)
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<AAAA>>', mes_ano[-4:])
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<MM>>', mes_ano[:2])
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<TIPO>>', v_diretorio_tipo)    

    log('DIRETÓRIO GIA............', v_diretorio )    
   
    for ie in v_lst_ies:
    
        v_mascara_arq_GIA = configuracoes.arq_gia
        v_mascara_arq_GIA = v_mascara_arq_GIA.replace('<<UF>>', uf)
        v_mascara_arq_GIA = v_mascara_arq_GIA.replace('<<MESANO>>', mes_ano)
        v_mascara_arq_GIA = v_mascara_arq_GIA.replace('<<IE>>', ie)
        v_mascara_arq_GIA = v_mascara_arq_GIA.replace('<<TIPO_ABREVIADO>>', v_mascara_tipo)
        v_mascara_arq_GIA = v_mascara_arq_GIA.replace('<<NNN>>', '*')
        
        log('MASCARA GIA .............', v_mascara_arq_GIA)

        arquivo = ultimo_Arquivo_Diretorio(v_mascara_arq_GIA, v_diretorio)
         
        log('ARQUIVO GIA .............', arquivo)
        if arquivo:
            log('Lendo arquivo ', arquivo)
            fd = open(os.path.join(v_diretorio,arquivo), 'r', encoding=tipoArquivo(os.path.join(v_diretorio,arquivo)))

            v_valor_apuracao = float(0.0)
            for linha in fd:

                if linha.startswith('07'):
                    v_valor_apuracao = float(linha[3:18])/100

                if linha.startswith('10'):

                    registro = {}
                    registro['Empresa']          = 'TBRA'
                    registro['UF Filial']        = uf
                    registro['Mês / Ano']        = mes_ano
                    registro['Série']            = 'N/a'
                    registro['CFOP']             = linha[2:6].upper().strip()
                    registro['Valor Líquido']    = float(linha[8:23])/100
                    registro['Valor Base']       = float(linha[23:38])/100
                    registro['Valor de ICMS']    = float(linha[38:53])/100
                    registro['Valor de Isentas'] = float(linha[53:68])/100
                    registro['Valor de Outras']  = float(linha[68:83])/100
                    #registro['VLR_IMPOSTO_RETIDO'] = float(linha[83:98])/100
                    registro['Substituo']        = float(linha[98:113])/100
                    registro['Substituído']      = float(linha[113:128])/100
                    registro['Outros Impostos']  = float(linha[128:143])/100
                    registro['Valor GIA']        = v_valor_apuracao
                
                    dados_retorno.append(registro)

            fd.close()

    return { 'dados' : dados_retorno,  'status' : 'Ok' }