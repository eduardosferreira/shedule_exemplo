#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
MODULO ...: TESHUVA
SCRIPT ...: retorna_dados_banco.py
CRIACAO ..: 13/01/2022
AUTOR ....: Victor Santos Cardoso - KYROS TECNOLOSPED
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
from openpyxl import Workbook, load_workbook
from pathlib import Path
from openpyxl.utils import get_column_letter
import fnmatch

def identifica_ies(uf, mes_ano, tipo) :

    if tipo == 'Atual Ti':  
        v_diretorio_tipo = 'ENXERTADOS'
        # - [PTITES-1634] 
        v_diretorio = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'ENXERTADOS', uf, mes_ano[-4:], mes_ano[:2])
    else:
        v_diretorio_tipo = 'PROTOCOLADOS'
        # - [PTITES-1634] 
        v_diretorio = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'SPED_FISCAL', 'PROTOCOLADOS', uf, mes_ano[-4:], mes_ano[:2])
    
    v_ies = []
    # - [PTITES-1634] -- v_diretorio = configuracoes.diretorio_sped
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<UF>>', uf)
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<AAAA>>', mes_ano[-4:])
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<MM>>', mes_ano[:2])
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<TIPO>>', v_diretorio_tipo)

    v_mascara_arq_SPED = configuracoes.arq_sped
    v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<MESANO>>', mes_ano)
    v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<UF>>', uf)
    v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<IE>>', '*')
    v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<TIPO_ABREVIADO>>', '*')
    v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<NNN>>', '*')

    l_diretorio = Path(v_diretorio)
    l_arq = l_diretorio.glob(v_mascara_arq_SPED)
    l_procura_arquivos = sorted(l_arq, reverse=False)
    if l_procura_arquivos:        
        for arquivo in l_procura_arquivos:
            ie = os.path.basename(str(arquivo)).split('_')[3]

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

def converter_valores(uf, mes_ano, tipo):
    
    dados_retorno = []
    if tipo == 'Atual Ti':  
        v_diretorio_tipo = 'ENXERTADOS'
        v_mascara_tipo = 'ENX*'
        # - [PTITES-1634] 
        v_diretorio = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'ENXERTADOS', uf, mes_ano[-4:], mes_ano[:2])
    else:
        v_diretorio_tipo = 'PROTOCOLADOS'
        v_mascara_tipo = 'PROT'
        # - [PTITES-1634] 
        v_diretorio = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'SPED_FISCAL', 'PROTOCOLADOS', uf, mes_ano[-4:], mes_ano[:2])
    
    v_lst_ies = identifica_ies( uf, mes_ano, tipo )
    # - [PTITES-1634] v_diretorio = configuracoes.diretorio_sped
    # - [PTITES-1634] v_diretorio = v_diretorio.replace('<<UF>>', uf)
    # - [PTITES-1634] v_diretorio = v_diretorio.replace('<<AAAA>>', mes_ano[-4:])
    # - [PTITES-1634] v_diretorio = v_diretorio.replace('<<MM>>', mes_ano[:2])
    # - [PTITES-1634] v_diretorio = v_diretorio.replace('<<TIPO>>', v_diretorio_tipo)    
    log('DIRETORIO SPED...........', v_diretorio) 
# para cada IE em v_lst_ies faça
    for ie in v_lst_ies:
        v_mascara_arq_SPED = configuracoes.arq_sped
        v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<UF>>', uf)
        v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<MESANO>>', mes_ano)
        v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<IE>>', ie)
        v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<TIPO_ABREVIADO>>', v_mascara_tipo)
        v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<NNN>>', '*')
        log('MASCARA SPED.............', v_mascara_arq_SPED) 

        arquivo = ultimo_Arquivo_Diretorio(v_mascara_arq_SPED, v_diretorio)

        log('ARQUIVO SPED ............', arquivo)         
        
        if arquivo:
    # abra o arquivo em modo leitura ('r')
            log('Lendo arquivo ', arquivo)

            fd = open(os.path.join(v_diretorio,arquivo), 'r', encoding=tipoArquivo(os.path.join(v_diretorio,arquivo)))

            v_E110           = 0
            v_E116           = 0
            v_serie          = "N/a"
            v_volume         = "N/a"
            v_nome           = "N/a"
            v_hash           = "N/a"
            v_cfop           = 0
            v_vlr_oper       = 0.0
            v_vlr_base_icms  = 0.0
            v_vlr_icms       = 0.0
            v_vlr_isentas    = 0.0
            v_vlr_outras     = 0.0
            v_vlr_outros_imp = 0.0
            v_vlr_aliquota   = 0.0
            v_cst            = 0
            v_ret            = "N/a"
            v_vlr_icms_st    = 0.0

            for linha in fd:
                v_dados  = linha.split("|")  
                if len(v_dados) < 7:
                    continue
                v_tp_reg = linha.split("|")[1]
                v_vlr_isentas = 0.0
                v_vlr_outras  = 0.0

                if v_tp_reg in ['C100','C500','D500','D100']:
                    
                    if v_dados[7] == '':
                        v_serie      = "N/a"
                    else:
                        v_serie      = v_dados[7]
                    
                    v_volume         = "N/a"
                    v_nome           = "N/a"
                    v_hash           = "N/a"
                    v_cfop           = 0
                    v_vlr_oper       = 0.0
                    v_vlr_base_icms  = 0.0
                    v_vlr_icms       = 0.0
                    v_vlr_isentas    = 0.0
                    v_vlr_outras     = 0.0
                    v_vlr_outros_imp = 0.0
                    v_vlr_aliquota   = 0.0
                    v_cst            = 0
                    v_ret            = "N/a"
                    v_vlr_icms_st    = 0.0
                
                if v_tp_reg == 'D695':
                    v_serie  = v_dados[3]  
                    v_volume = v_dados[8].strip()[-3:]
                    v_nome   = v_dados[8].strip()
                    v_hash   = v_dados[9].strip()
                    
                    if mes_ano[2:] >= '2017':                    
                        v_ret = v_dados[8].strip()[-8]
                    else:
                        v_ret = v_dados[8].strip()[-6]
                                    
                if v_tp_reg in ['C190','C590','D590','D696','D190']:
                    v_cfop          = v_dados[3].strip()
                    v_vlr_oper      = float(v_dados[5].replace(',','.'))
                    v_vlr_base_icms = float(v_dados[6].replace(',','.'))
                    v_vlr_icms      = float(v_dados[7].replace(',','.'))
                    v_cst           = v_dados[2]
                    
                    if v_dados[4]:    
                        v_vlr_aliquota  = float(v_dados[4].replace(',','.'))
                    else:
                        v_vlr_aliquota  = 0.0

                    if v_tp_reg == 'C190': 
                        v_vlr_outros_imp = float(v_dados[11].replace(',','.')) 
                        v_vlr_icms_st = float(v_dados[9].replace(',','.'))

                    if v_cst.endswith('30') or v_cst.endswith('40') or v_cst.endswith('41'):
                        v_vlr_isentas = v_vlr_oper

                    if v_cst.endswith('20'):
                        v_vlr_isentas = v_vlr_oper - v_vlr_base_icms

                    if v_cst.endswith('50') or v_cst.endswith('51') or v_cst.endswith('60') or v_cst.endswith('70') or v_cst.endswith('90'):
                        v_vlr_outras = v_vlr_oper                    

                    registro = {}
                    registro['Arquivo']                                  = arquivo
                    registro['Empresa']                                  = 'TBRA'
                    registro['UF Filial']                                = uf
                    registro['Mês / Ano']                                = mes_ano
                    registro['Série']                                    = v_serie
                    registro['Volume']                                   = v_volume.lstrip('0')
                    registro['CFOP']                                     = v_cfop
                    registro['Valor Líquido']                            = v_vlr_oper
                    registro['Valor Base']                               = v_vlr_base_icms
                    registro['Valor de ICMS']                            = v_vlr_icms
                    registro['Valor de Isentas']                         = v_vlr_isentas
                    registro['Valor de Outras']                          = v_vlr_outras
                    registro['Outros Impostos']                          = v_vlr_outros_imp
                    registro['Alíquota']                                 = v_vlr_aliquota
                    registro['CST']                                      = v_cst
                    registro['Nome do Arquivo Mestre']                   = v_nome
                    registro['Código de Autenticação']                   = v_hash
                    registro['Indicador Retificação']                    = v_ret
                    registro['Substituo']                                = v_vlr_icms_st
               
                    dados_retorno.append(registro)

                if v_tp_reg == 'E110':
                    v_E110 += float(v_dados[13].replace(',','.'))
                
                if v_tp_reg == 'E116':
                    if v_dados[5] == '046-2':
                        v_E116 += float(v_dados[3].replace(',','.'))
            fd.close()

            for registro in dados_retorno:
                if registro['Arquivo'] == arquivo: 
                    registro['E110'] = v_E110
                    registro['E116'] = v_E116
    
    return { 'dados' : retira_duplicidade_sped(dados_retorno),  'status' : 'Ok' }

def retira_duplicidade_sped(p_registros):

    retorno = []

    ix_linha = 0
# ordernar p_registros por 'Arquivo', 'Empresa', 'UF Filial', 'Mês / Ano', 'Série', 'Volume', 'Alíquota', 'CST' e 'CFOP'
    p_registros = sorted(p_registros, key=lambda row:(row['Empresa'],row['UF Filial'],row['Mês / Ano'],row['Série'],row['Volume'],row.get('Alíquota','N/a'),row.get('CST','N/a'),row['CFOP']),reverse=False)
    
    v_chave_linha_anterior = ''
# Para v_linha existentes em p_registros:
    ix_linha= ix_linha+ 1

    for v_linha in p_registros:

        v_chave_linha = str(v_linha['Arquivo']).strip()   +'|'+\
                        str(v_linha['Empresa']).strip()   +'|'+\
                        str(v_linha['UF Filial']).strip() +'|'+\
                        str(v_linha['Mês / Ano']).strip() +'|'+\
                        str(v_linha['Série']).strip()     +'|'+\
                        str(v_linha['Volume']).strip()    +'|'+\
                        str(v_linha['Alíquota']).strip()  +'|'+\
                        str(v_linha['CST']).strip()       +'|'+\
                        str(v_linha['CFOP']).strip()      +'|' 

# Se v_chave_linha_anterior = v_chave_linha
        if v_chave_linha_anterior == v_chave_linha:
            
            retorno[-1]['Valor Líquido']    = retorno[-1]['Valor Líquido']    + v_linha['Valor Líquido']
            retorno[-1]['Valor Base']       = retorno[-1]['Valor Base']       + v_linha['Valor Base']
            retorno[-1]['Valor de ICMS']    = retorno[-1]['Valor de ICMS']    + v_linha['Valor de ICMS']
            retorno[-1]['Valor de ISENTAS'] = retorno[-1]['Valor de Isentas'] + v_linha['Valor de Isentas']
            retorno[-1]['Valor de Outras']  = retorno[-1]['Valor de Outras']  + v_linha['Valor de Outras']
            retorno[-1]['Outros Impostos']  = retorno[-1]['Outros Impostos']  + v_linha['Outros Impostos']
# Senão
        else:
            retorno.append(v_linha)
            v_chave_linha_anterior = v_chave_linha

    return retorno
    

def converter_nome_arquivos_d695(uf, mes_ano, tipo):

    log('converter_nome_arquivos_d695')

    dados_retorno = []
    if tipo == 'Atual Ti':  
        v_diretorio_tipo = 'ENXERTADOS'
        v_mascara_tipo = 'ENX*'
        # - [PTITES-1634] 
        v_diretorio = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'ENXERTADOS', uf, mes_ano[-4:], mes_ano[:2])
    else:
        v_diretorio_tipo = 'PROTOCOLADOS'
        v_mascara_tipo = 'PROT'
        # - [PTITES-1634] 
        v_diretorio = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'SPED_FISCAL', 'PROTOCOLADOS', uf, mes_ano[-4:], mes_ano[:2])
 
    v_lst_ies = identifica_ies( uf, mes_ano, tipo )
    # - [PTITES-1634] -- v_diretorio = configuracoes.diretorio_sped
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<UF>>', uf)
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<AAAA>>', mes_ano[-4:])
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<MM>>', mes_ano[:2])
    # - [PTITES-1634] -- v_diretorio = v_diretorio.replace('<<TIPO>>', v_diretorio_tipo)    

    log('DIRETORIO ABA 3 ', v_diretorio)
        
    for ie in v_lst_ies:
        v_mascara_arq_SPED = configuracoes.arq_sped
        v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<UF>>', uf)
        v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<MESANO>>', mes_ano)
        v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<IE>>', ie)
        v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<TIPO_ABREVIADO>>', v_mascara_tipo)
        v_mascara_arq_SPED = v_mascara_arq_SPED.replace('<<NNN>>', '*')
        arquivo = ultimo_Arquivo_Diretorio(v_mascara_arq_SPED, v_diretorio)

        log('MASCARA ABA 3 ', v_mascara_arq_SPED)
        log('ARQUIVO ABA 3 ', arquivo)

        if arquivo:
            fd = open(os.path.join(v_diretorio,arquivo), 'r', encoding='iso-8859-1')

            for linha in fd:
                v_tp_reg = linha.split('|')[1]

                if v_tp_reg == "D695":
                    v_dados = linha.split('|')
                    v_serie = v_dados[3]
                    v_volume = v_dados[8].strip()[-3:]
                    v_nome = v_dados[8].strip()
                    v_hash = v_dados[9].strip()

                    registro = {}
                    registro['Arquivo']                     = arquivo
                    registro['Empresa']                     = 'TBRA'
                    registro['UF Filial']                   = uf
                    registro['Mês / Ano']                   = mes_ano
                    registro['Série']                       = v_serie
                    registro['Volume']                      = v_volume
                    registro['SPED - ' + tipo + ' - Nome']  = v_nome
                    registro['SPED - ' + tipo + ' - Hash']  = v_hash
                    
                    dados_retorno.append(registro)

            fd.close()
    return { 'dados' : dados_retorno,  'status' : 'Ok' }