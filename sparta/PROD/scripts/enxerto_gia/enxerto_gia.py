#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: GF
  MODULO ...: 
  SCRIPT ...: enxerto_gia.py
  CRIACAO ..: 14/11/2019
  AUTOR ....: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO : VARRE O DIRETORIO DE ARQUIVOS PROCURANDO ARQUIVOS COM A MASCARA 
              '*GIA.csv' E PARA CADA ARQUIVO EXECUTA O PROCEDIMENTO DESCRITO NA ESPECIFICACAO
              Teshuva_EspecificacaoFuncionalEnxertoGIA_v1.1.docx
              
----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 14/11/2019 - Welber Pena de Sousa - Kyros Tecnologia
            - Criacao do script.

    * 04/12/2019 - Welber Pena de Sousa - Kyros Tecnologia
            - Alteracao do script para atender o novo documento 
                Teshuva_EspecificacaoFuncionalEnxertoGIA_v1.2.docx 
                    (Alteracao : RF002 - Enxerto DIPAM)

    * 24/07/2020 - Welber Pena de Sousa - Kyros Tecnologia
            - Alteração do script para utilizar os diretorios como citado no documento
                Teshuva_EspecificaçãoFuncional_Enxerto_GIA_v2.docx
                    (Alteração : ALT001 - Diretorios de arquivos)
                    
    * 05/02/2021 - Airton Borges da Silva - Kyros Tecnologia
        - Alteração do script para receber parametros UF, IE e MES,ANO. 
        - Alteração os diretorios como citado no documento
            "12 - Fase 2 - Alteração Enxerto GIA"
    
    - 04/03/2021 - - Airton Borges da Silva - Kyros Tecnologia
        - Alterado de AAAAMM para MMAAAA os nomes dos arquivos. 

    - 23/08/2021  - Marcelo Gremonesi - Kyros Tecnologia
        -   Adquacoes para o novo painel
    
    - 18/11/2021 - Fabrisia G. Rosa - Kyros Tecnologia
        - Alterações no script. PTITES-1029 - DV - Alteração Enxerto GIA - Calculo DIPAM - Codificação 

    - 22/02/2022 - Eduardo da Silva Ferreira - Kyros Tecnologia
            - [PTITES-1634] Padrão de diretórios do SPARTA

----------------------------------------------------------------------------------------------
"""

import sys
import os

SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)
import configuracoes

import cx_Oracle
import unicodedata
import fnmatch
import shutil
import datetime
import atexit
import re

from pathlib import Path
#from comum import *

import comum
import sql
import util

#import databases
#import config

ret = 0
nome_relatorio = "" 
dir_destino = "" 
dir_dados = "" 
iei = ""
datref = ""

disco = ('' if os.name == 'posix' else 'D:')
#SD = ('/' if os.name == 'posix' else '\\')
name_script = os.path.basename(__file__).split('.')[0]


dic_registros = {}
dic_layouts = {}
dic_campos = {}
dic_indicesDIPAM = {}
# dic_registros['register'] = {}
# dic_registros['trailer'] = {}

##### Estrutura do tipo de registro 01 – Registro Mestre.
dic_registros['01'] = {}
dic_registros['01'][0] = [ "CR", 2 ]
dic_registros['01'][1] = [ "TipoDocto", 2 ]
dic_registros['01'][2] = [ "DataGeracao", 8 ]
dic_registros['01'][3] = [ "HoraGeracao", 6 ]
dic_registros['01'][4] = [ "VersaoFrontEnd", 4 ]
dic_registros['01'][5] = [ "VersaoPref", 4 ]
dic_registros['01'][6] = [ "Q05", 4 ]
dic_registros['01'][7] = [ "FimLinha", 2 ]

##### Estrutura do tipo de registro 05 – Cabeçalho do Documento Fiscal.
dic_registros['05'] = {}
dic_registros['05'][0] = [ 'CR', 2 ]
dic_registros['05'][1] = [ 'IE', 12 ]
dic_registros['05'][2] = [ 'CNPJ', 14 ]
dic_registros['05'][3] = [ 'CNAE', 7 ]
dic_registros['05'][4] = [ 'RegTrib', 2 ]
dic_registros['05'][5] = [ 'Ref', 6 ]
dic_registros['05'][6] = [ 'RefInicial', 6 ]
dic_registros['05'][7] = [ 'Tipo', 2 ]
dic_registros['05'][8] = [ 'Movimento', 1 ]
dic_registros['05'][9] = [ 'Transmitida', 1 ]
dic_registros['05'][10] = [ 'SaldoCredPeriodoAnt', 15 ]
dic_registros['05'][11] = [ 'SaldoCredPeriodoAntST', 15 ]
dic_registros['05'][12] = [ 'OrigemSoftware', 14 ]
dic_registros['05'][13] = [ 'OrigemPreDig', 1 ]
dic_registros['05'][14] = [ 'ICMSFixPer', 15 ]
dic_registros['05'][15] = [ 'ChaveInterna', 32 ]
dic_registros['05'][16] = [ 'Q07', 4 ]
dic_registros['05'][17] = [ 'Q10', 4 ]
dic_registros['05'][18] = [ 'Q20', 4 ]
dic_registros['05'][19] = [ 'Q30', 4 ]
dic_registros['05'][20] = [ 'Q31', 4 ]
dic_registros['05'][21] = [ 'FimLinha', 2 ]

##### Estrutura do tipo de registro 10 – Detalhes CFOPs.
dic_registros['10'] = {}
dic_registros['10'][0] = [ 'CR', 2 ]
dic_registros['10'][1] = [ 'CFOP', 6 ]
dic_registros['10'][2] = [ 'ValorContabil', 15 ]
dic_registros['10'][3] = [ 'BaseCalculo', 15 ]
dic_registros['10'][4] = [ 'Imposto', 15 ]
dic_registros['10'][5] = [ 'IsentasNaoTrib', 15 ]
dic_registros['10'][6] = [ 'Outras', 15 ]
dic_registros['10'][7] = [ 'ImpostoRetidoST', 15 ]
dic_registros['10'][8] = [ 'ImpRetSubstitutoST', 15 ]
dic_registros['10'][9] = [ 'ImpRetSubstituido', 15 ]
dic_registros['10'][10] = [ 'OutrosImpostos', 15 ]
dic_registros['10'][11] = [ 'Q14', 4 ]
dic_registros['10'][12] = [ 'FimLinha', 2 ]

##### Estrutura do tipo de registro 14 – Detalhes Interestaduais.
dic_registros['14'] = {}
dic_registros['14'][0] = [ 'CR', 2 ]
dic_registros['14'][1] = [ 'UF', 2 ]
dic_registros['14'][2] = [ 'Valor_Contabil_1', 15 ]
dic_registros['14'][3] = [ 'BaseCalculo_1', 15 ]
dic_registros['14'][4] = [ 'Valor_Contabil_2', 15 ]
dic_registros['14'][5] = [ 'BaseCalculo_2', 15 ]
dic_registros['14'][6] = [ 'Imposto', 15 ]
dic_registros['14'][7] = [ 'Outras', 15 ]
dic_registros['14'][8] = [ 'ICMSCobradoST', 15 ]
dic_registros['14'][9] = [ 'PetroleoEnergia', 15 ]
dic_registros['14'][10] = [ 'Outros Produtos', 15 ]
dic_registros['14'][11] = [ 'Benef', 1 ]
dic_registros['14'][12] = [ 'Q18', 4 ]
dic_registros['14'][13] = [ 'FimLinha', 2 ]

##### Estrutura do tipo de registro 14 – Detalhes Interestaduais.
dic_registros['30'] = {}
dic_registros['30'][0] = [ 'CR', 2 ]
dic_registros['30'][1] = [ 'CodDIP', 2 ]
dic_registros['30'][2] = [ 'Municipio', 5 ]
dic_registros['30'][3] = [ 'Valor', 15 ]
dic_registros['30'][4] = [ 'FimLinha', 2 ]

def busca_valor_contabil(ie, mes, ano, uf):
    """
        PTITES-1029 - Função que busca o calculo do DIPAM dos CFOP's 1301, 2301 e 3301 
        à partir do arquivo do SPED PROTOCOLADO.
    """

    mask_arq_sped_fiscal_prot = configuracoes.arq_sped_fiscal_protocolado.replace("<<MESANO>>", '%s%s'%(mes, ano)).replace("<<UF>>", uf).replace("<<IE>>", ie).replace("<<NNN>>","*").strip()
    # Inicio [PTITES-1634]
    # dir_sped_prot = configuracoes.sped_fiscal_protocolado.replace("<<MM>>", mes).replace("<<AAAA>>", ano).replace("<<UF>>", uf).strip()
    dir_sped_prot = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'SPED_FISCAL', 'PROTOCOLADOS', uf, ano, mes)
    arquivo = nome_arquivo(mask_arq_sped_fiscal_prot, dir_sped_prot)
    try:
        file_exists = os.path.exists(arquivo)
    except:
        file_exists = False
        log("Nao existe o arquivo ..", mask_arq_sped_fiscal_prot, dir_sped_prot)
        return None
    # Fim [PTITES-1634]
    
    # inicializa as variaveis de controle
    fl_sair = 0
    contador = 0
    nregrel = 0
    tp_reg = ""

    cfop = ""
    vlr_oper = float(0)

    log("Busncando valor contábil..")

    if file_exists: #[PTITES-1634] arquivo.is_file():
        log(" - Processando leitura do arquivo.: " , arquivo)

        #abre o arquivo
        ent = open(arquivo, mode='r', encoding=encodingDoArquivo(arquivo))

        #lê as linhas do arquivo
        linha_lida = ent.readline()

        #percorre o arquivo
        while (linha_lida and fl_sair == 0):
            contador += 1
            #quebra em vetor
            dados_entr = linha_lida.split("|")
            
            tp_reg = ""
            if (len(dados_entr) >= 7):
                tp_reg = dados_entr[1].upper().strip()
            
            #valida o tipo
            if tp_reg == "D590":
                nregrel += 1

                cfop = dados_entr[3].upper().strip()
                if int(cfop) in [ 1301, 2301, 3301]:
                    vlr_oper += float(dados_entr[5].replace(',','.'))
            
            linha_lida = ent.readline()
        
        #fecha o arquivo
        ent.close()

        #verifica se houve registros processados
        if nregrel > 0:
            return vlr_oper
        else:
            log("# " + str(len(nregrel)) + " >> " + "Não processou nenhum dado do arquivo.: " + arquivo)
            return None
    
    else:
        return None






class Relatorio:
    def __init__(self, nome_relatorio, dir_geracao = ''):
        self.__arquivo_relatorio = '%s.csv'%( nome_relatorio )
        if not dir_geracao :
            self.__dir_geracao = '.'
        else :
            self.__dir_geracao = dir_geracao
        self.__path_relatorio = os.path.join(self.__dir_geracao, self.__arquivo_relatorio)

        self.__colunas_relatorio = []

        self.__linhas_relatorio = []
        atexit.register(self.close)

    @property
    def diretorio_geracao(self): return self.__dir_geracao
    @property
    def arquivo_relatorio(self): return self.__arquivo_relatorio
    @property
    def qtd_linhas_relatorio(self): return len(self.__linhas_relatorio)

    def adicionar_coluna(self, coluna):
        self.__colunas_relatorio.append(coluna)

    def registrar(self, *args):
        if len(args) == len(self.__colunas_relatorio) :
            # print('Registrar :', args)
            self.__linhas_relatorio.append([x for x in args])
        else :
            log("Erro ao registar dados no relatorio.")
            log(*args)
            raise
        return True
        

    def close(self):
        if len(self.__linhas_relatorio) > 0 :
            fd = open( self.__path_relatorio, 'w' )
            fd.write(';'.join( x.replace('\n', '') for x in self.__colunas_relatorio ) + '\n')
            for linha in self.__linhas_relatorio :
                fd.write( ';'.join( str(x).replace('\n', '').replace('\r', '') for x in linha ) + '\n' )
            fd.close()
            log('Gerado relatorio ..: %s'%(self.__arquivo_relatorio))


def encodingDoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r', encoding='iso-8859-1')
        t = fd.read()
        fd.close()
    except :
        return 'utf-8'

    return 'iso-8859-1'

def criarDicionarioLayouts() :
    log('Gerando dicionario de Layouts.')
    for key in dic_registros.keys() :
        p = 0
        dic_layouts[key] = {}
        dic_campos[key] = {}
        for y in range(len(dic_registros[key].keys())) :
            field = dic_registros[key][y]
            dic_layouts[key][y] = [ field[0], p ]
            p += field[1]
            dic_layouts[key][y].append(p) 
            dic_campos[key][field[0]] = y
    log('Dicionario de Layouts criado !')
    return True

def leIndiceMunicipioDIPAM(ano, mes, uf) :
    """
        Função que busca os dados à partir da tabela gfcadastro.reg_1400_municipio_sp. PTITES-1029
    """

    log('Gerando dicionario de Indices DIPAM')
    try :
        query = """
                SELECT g.COD_MUN_GIA AS Codigo
                      ,g.MIBGE_DESC_MUN AS Municipio
                      ,m.INDICE
                FROM gfcadastro.reg_1400_municipio_sp m
                INNER JOIN gfcadastro.REG_1400_MUNICIPIO_DE_PARA_GIA g  ON g.MIBGE_COD_MUN = m.CODIGO_IBGE
                WHERE 1=1
                AND m.ano = '%s'  --PARAMETRO
                AND m.mes = '%s'  --PARAMETRO
                AND m.uf  = '%s'  --PARAMETRO
                AND m.data_fim IS NULL
                ORDER BY g.MIBGE_DESC_MUN
          
             """%(ano, mes, uf)
        
        banco.executa(query)
        result = banco.fetchall()

        if (result != None):
            for item in result:
                dic_indicesDIPAM[int(item[0])] = [ item[1], item[2]]
    except:
        log('ERRO - Impossivel criar o dicionario de Indices DIPAM')
        return False

    return True

def cfopDeEntrada(cfop) :
    if int(cfop) in ( 1301, 2301, 3301 ) :
        return True
    return False

def cfopDeveSerConsiderado(cfop) :
    if int(cfop) >= 5300 and int(cfop) <= 5307 :
        return True
    if int(cfop) >= 6300 and int(cfop) <= 6307 :
        return True
    if int(cfop) >= 7300 and int(cfop) <= 7301 :
        return True
    return False

def quebraRegistro(reg) :
    cr = reg[:2]
    itens_registro = []
    colunas = []
    for y in range( len(dic_layouts[cr].keys()) ) :
        field, i, f = dic_layouts[cr][y]
        itens_registro.append( reg[i:f] )
        colunas.append(field)
    # print(colunas)
    return itens_registro

def processaDiretorio(dnprot, dnreg, nenx, nrerr, denx) :
    global relatorio_erros 
    global iei
    global datref
    global ufi


    if not os.path.isdir(denx) :
        os.makedirs(denx)

    relatorio_erros = Relatorio( 'ERR_%s'%(nrerr), denx )
    relatorio_erros.adicionar_coluna('Tipo')
    relatorio_erros.adicionar_coluna('Mensagem')
    
    dnenx      = os.path.join(denx,nenx) 
    encodingp   = encodingDoArquivo(dnprot) 
    encodingr   = encodingDoArquivo(dnreg)
    obj_dnreg  = None

    obj_dnenx  = open(dnenx, 'w', encoding = 'iso-8859-1')
    obj_dnprot = open(dnprot, 'r', encoding = encodingp )

    log('Arquivo a ser gerado ..: %s'%(dnenx))

    x = 0
    lin = obj_dnprot.readline()
    novas_linhas = []
    valorEntradas = 0
    valorSaidas = 0
    enxertou = False
    path_arq = dnprot
  
    while lin :
    # for lin in obj_dnprot :
        
        if lin.startswith('05') :
            log('Encontrado novo documento .')
            registro = quebraRegistro(lin)
            cabecalho_corrente = registro[:]
            linhas_a_escrever = []
            ### Busca o documento '05' referente no arquivo regerado.
            if obj_dnreg :
                obj_dnreg.close()
                del obj_dnreg
            
            #### Busca o arquivo REGERADO na pasta de arquivos regerados.
            obj_dnreg = open(dnreg, 'r', encoding = encodingr )
            achou = 0
            registro_regerado = [7]
            while achou < 1 :
                l = obj_dnreg.readline()
                if l.startswith('05') :
                    registro_regerado = quebraRegistro(l)
                    # print(registro_regerado)
                    if registro_regerado[:6] == registro[:6] :
                        achou = 1
                if not l :
                    achou = 2
            novas_linhas = []
            if achou == 1 :
                fim_documento = False
                while not fim_documento :
                    prox_lin_regerado = obj_dnreg.readline()
                    if not prox_lin_regerado.startswith('05') :
                        if prox_lin_regerado :
                            #########################################################################
                            #### Alteracao : Welber Pena - 04/12/2019
                            if prox_lin_regerado[:2] in ( '10', '14', '18' ) : ### , '30' ) :
                            ####
                                novas_linhas.append(prox_lin_regerado)
                                if prox_lin_regerado.startswith('10') :
                                    r10 = quebraRegistro(prox_lin_regerado)
                                    idx_cfop = dic_campos['10']['CFOP']
                                    cfop = r10[idx_cfop][:4]
                                    #########################################################################
                                    #### Alteracao : Welber Pena - 04/12/2019
                                    if cfopDeveSerConsiderado(cfop) :
                                        idx_ValorContabil = dic_campos['10']['ValorContabil']
                                        valorSaidas += int(r10[idx_ValorContabil].lstrip('0') or '0')
                                    ####
                        else :
                            fim_documento = True
                    else :
                        fim_documento = True
            #### Encontrou as linhas referentes no arquivo regerado.
            if novas_linhas :
                enxertou = True

            #### Pega todas as linhas referentes ao documento.
            log('Buscando todas as linhas referentes ao documento.')
            linhas_protocolado = []
            fim_documento = False
            while not fim_documento :
                lin = obj_dnprot.readline()
                if not lin.startswith('05') :
                    if lin :
                        linhas_protocolado.append(lin)
                        if lin.startswith('10') :
                            r10 = quebraRegistro(lin)
                            idx_cfop = dic_campos['10']['CFOP']
                            cfop = r10[idx_cfop][:4]
                            #########################################################################
                            #### Alteracao : Welber Pena - 04/12/2019
                            if cfopDeEntrada(cfop) :
                                idx_ValorContabil = dic_campos['10']['ValorContabil']
#                                print("  = ", )
#                                print("valorEntradas  = ",valorEntradas )
#                                print("idx_ValorContabil  = ",idx_ValorContabil )
#                                print("dic_campos  = ", dic_campos)
#                                print("dic_campos['10'] = ",dic_campos['10'] )
#                                print("r10[idx_ValorContabil].lstrip('0')  = ",r10[idx_ValorContabil] )
#                                print("  = ", )
#                                print("  = ", )
#                               
#                                valorEntradas += int(r10[idx_ValorContabil])
                                
                                valorEntradas += int( '0' + str(r10[idx_ValorContabil].lstrip('0')))
                                
#                                valorEntradas += int(r10[idx_ValorContabil].lstrip('0'))
                            ####
                    else :
                        fim_documento = True
                else :
                    fim_documento = True
            #########################################################################
            #### Alteracao : Welber Pena - 04/12/2019
            valorEntradas = (float(busca_valor_contabil(iei, datref[2:4], datref[-4:], ufi )) * 100)   #PTITES-1029 
            log('Valor Entradas : %s'%( valorEntradas ))
            log('Valor Saidas   : %s'%( valorSaidas ))
            fator = (valorSaidas - valorEntradas ) /100.0
            log('Valor Fator    : %s'%( fator ))
            ##### 
        elif lin.startswith('01') :
            log('Gerando HEADER do arquivo enxertado.')
            obj_dnenx.write(lin.replace('\n', '\r\n'))
            lin = obj_dnprot.readline()
        if novas_linhas and linhas_protocolado and not linhas_a_escrever :
            q10 = 0
            log('Escrevendo registros do documento.')
            #### Retiro todas as linhas '10' e seus filhos '14' e '18' do protocolado.
            while len(linhas_protocolado) > 0 :
                l = linhas_protocolado.pop(0)
                if l.startswith('10') :
                    r10 = quebraRegistro(l)
                    idx_cfop = dic_campos['10']['CFOP']
                    cfop = r10[idx_cfop][:4]
                    # print('CFOP :', cfop)
                    if cfopDeveSerConsiderado(cfop[:4]) :
                        #### Retiro os filhos.
                        # print('Retirando 10 e filhos ....')
                        retirar_filhos = True
                        while retirar_filhos and len(linhas_protocolado) > 0 :
                            l = linhas_protocolado[0]
                            if l.startswith('14') or l.startswith('18') :
                                l = linhas_protocolado.pop(0)
                            else :
                                retirar_filhos = False
                
                    existe_cfop_menor = True
                    y_novas_linhas = 0
                    while existe_cfop_menor :
                        new_line = novas_linhas[y_novas_linhas]
                        if new_line.startswith('10') :
                            t = quebraRegistro(new_line)
                            t_cfop = t[idx_cfop][:4]
                            if cfopDeveSerConsiderado(t_cfop) :
                                if int(t_cfop) <= int(cfop) :
                                    linhas_a_escrever.append(novas_linhas.pop(y_novas_linhas))
                                    q10 += 1
                                    inserir_filhos = True
                                    while inserir_filhos and len(novas_linhas) > 0 and y_novas_linhas < len(novas_linhas) :
                                        new_line = novas_linhas[y_novas_linhas]
                                        if new_line.startswith('14') or new_line.startswith('18') :
                                            linhas_a_escrever.append(novas_linhas.pop(y_novas_linhas))
                                        else :
                                            inserir_filhos = False
                                    if y_novas_linhas > 0 :
                                        y_novas_linhas -= 1
                        y_novas_linhas += 1
                        if y_novas_linhas >= len(novas_linhas) :
                            existe_cfop_menor = False
                    
                    if not cfopDeveSerConsiderado(cfop[:4]) :
                        q10 += 1
                        linhas_a_escrever.append(l)

                elif int(l[:2]) > 18 and l[:2] != '30' :
                    if len(novas_linhas) > 0 :
                        existe_cfop_menor = True
                        y_novas_linhas = 0
                        while existe_cfop_menor :
                            new_line = novas_linhas[y_novas_linhas]
                            if new_line.startswith('10') :
                                t = quebraRegistro(new_line)
                                t_cfop = t[idx_cfop][:4]
                                if cfopDeveSerConsiderado(t_cfop) :
                                    linhas_a_escrever.append(novas_linhas.pop(y_novas_linhas))
                                    q10 += 1
                                    inserir_filhos = True
                                    while inserir_filhos and len(novas_linhas) > 0 and y_novas_linhas < len(novas_linhas) :
                                        new_line = novas_linhas[y_novas_linhas]
                                        if new_line.startswith('14') or new_line.startswith('18') :
                                            linhas_a_escrever.append(novas_linhas.pop(y_novas_linhas))
                                        else :
                                            inserir_filhos = False
                            y_novas_linhas += 1
                            if y_novas_linhas >= len(novas_linhas) :
                                existe_cfop_menor = False
                    linhas_a_escrever.append(l)
                elif l[:2] == '30' :
                    #########################################################################
                    # if len(linhas_protocolado) > 0 :
                    #     existe_tipo_30 = True
                    #     y_linhas_protocolado = 0
                    #     #### Retira todas as linhas do tipo '30'
                    #     while existe_tipo_30 :
                    #         new_line = linhas_protocolado[y_linhas_protocolado]
                    #         if new_line.startswith('30') :
                    #             l = linhas_protocolado.pop(0)
                    #         else :
                    #             y_linhas_protocolado += 1
                    #         if y_linhas_protocolado >= len(linhas_protocolado) :
                    #             existe_tipo_30 = False

                    # ##### Acrescenta todas as novas linhas do tipo '30' .
                    # if len(novas_linhas) > 0 :
                    #     y_novas_linhas = 0
                    #     while y_novas_linhas < len(novas_linhas) :
                    #         new_line = novas_linhas[y_novas_linhas]
                    #         if new_line.startswith('30') :
                    #             l = novas_linhas.pop(y_novas_linhas)
                    #             linhas_a_escrever.append(l)
                    #         else :
                    #             y_novas_linhas += 1
                    #########################################################################
                    #### Alteracao : Welber Pena - 04/12/2019
                    #### Alterada a regra acima onde os registros do tipo 30  do PROTOCOLADO
                    #### eram desprezados e os novos registros do REGERADO eram considerados.
                    ####
                    #### Agora ficou o inverso os registros do PROTOCOLADO serao mantidos
                    #### mas com a atualizacao dos seus valores.
                    #### Linha 256  
                    ####     >> if prox_lin_regerado[:2] in ( '10', '14', '18' ) : ### , '30' ) :
                    #### tambem foi alterada para isso .
                    #########################################################################
                    if len(linhas_protocolado) > 0 :
                        existe_tipo_30 = True
                        y_linhas_protocolado = 0
                        linhas_protocolado.insert(0, l)
                        #### Alterar os valores do registro tipo '30'
                        while existe_tipo_30 :
                            new_line = linhas_protocolado[y_linhas_protocolado]
                            if new_line.startswith('30') :
                                r30 = quebraRegistro(new_line)
                                idx_municipio = dic_campos['30']['Municipio']
                                idx_valor = dic_campos['30']['Valor']
                                # print(r30)
                                try :
                                    municipio, indice = dic_indicesDIPAM[int(r30[idx_municipio])]
                                except :
                                    log('ERRO - Codigo de municipio invalido ! < %s >'%(r30[idx_municipio]))
                                    relatorio_erros.registrar("ERRO", 'Codigo de municipio invalido ! < %s >'%(r30[idx_municipio]))
                                    return False
                                
                                indice = float( indice )
                                valor_DIPAM = (fator * indice)
                                r30[idx_valor] = str('%016.2f'%(round(valor_DIPAM, 2))).replace('.','')
                                l = linhas_protocolado.pop(y_linhas_protocolado)
                                linhas_a_escrever.append( ''.join( x for x in r30) )
#                                log('Recalculado valor para o municipio %s - %s - %016.2f >'%( municipio, r30[idx_municipio], round(valor_DIPAM, 2) ))
#                                log('  - Indice : %s | Fator : %s | Valor : %s'%(indice, fator, valor_DIPAM)) 
                            else :
                                y_linhas_protocolado += 1
                            if y_linhas_protocolado >= len(linhas_protocolado) :
                                existe_tipo_30 = False

                    #########################################################################
                else :
                    linhas_a_escrever.append(l)
            
            ##### Insiro todas as linhas '10' com CFOPs validos e seus filhos '14' e '18'
            while len(novas_linhas) > 0 :
                l = novas_linhas.pop(0)
                if l.startswith('10') :
                    r10 = quebraRegistro(l)
                    idx_cfop = dic_campos['10']['CFOP']
                    cfop = r10[idx_cfop][:4]
                    if cfopDeveSerConsiderado(cfop) :
                        ### Insiro a linha '10' e seus filhos '14' e '18'
                        linhas_a_escrever.append(l)
                        q10 += 1
                        inserir_filhos = True
                        while inserir_filhos and len(novas_linhas) > 0 and y_novas_linhas < len(novas_linhas) :
                            l = novas_linhas[0]
                            if l.startswith('14') or l.startswith('18') :
                                l = novas_linhas.pop(0)
                                linhas_a_escrever.append(l)
                            else :
                                inserir_filhos = False

            ##### Escrever o documento '05' e seus filhos no novo arquivo.
            if len(linhas_a_escrever) > 0 :
                ##### Acerto a quantidade de registros '10' no registro '05' do documento.
                idx_q10 = dic_campos['05']['Q10']
                cabecalho_corrente[idx_q10] = str(q10).rjust(4,'0')
                obj_dnenx.write( "".join(x for x in cabecalho_corrente).replace('\n', '\r\n') )
                for l in linhas_a_escrever :
                    obj_dnenx.write( l.replace('\n', '\r\n') )
                    x += 1
    # print(lin)

    if enxertou :
        log('Arquivo processado !')
        obj_dnprot.close()
        obj_dnenx.close()

        log(" R E S U M O ".center(50,'*'))
        log('** Numero de linhas processadas ..: %s'%(x))
        # log('** Qtde de doc nao encontrados ...: %s'%(linha_doc_nao_encontrado))
        # log('** Qtde de registros alterados ...: %s'%(atualizados))
        # log('** Qtde de reg ja alterados ......: %s'%(qtd_ja_alterados))
        # log('** Qtde de cad_cli inseridos .....: %s'%(qtd_cli_fornec_inseridos))
        # log('** Qtde de complvu inseridos .....: %s'%(qtd_complvu_inseridos))
        log('*'*50)

        obj_dnenx = open(dnenx, 'r', encoding = 'iso-8859-1')
        ultimas_linhas = ['','']
        for lin in obj_dnenx.readlines() :
            ultimas_linhas[1] = ultimas_linhas[0]
            ultimas_linhas[0] = lin
        obj_dnenx.close()

        ultimo_reg_30 = True
        if not ultimas_linhas[0].startswith('30') :
            if ultimas_linhas[0] ==  '' :
                if not ultimas_linhas[1].startswith('30') :
                    ultimo_reg_30 = False
            else :
                ultimo_reg_30 = False
        
        if not ultimo_reg_30 :
            if os.path.isfile(dnenx) :
                os.remove(dnenx)
            log("ERRO - Arquivo 'Enxertado' não possui a ultima linha com registro 30")
            relatorio_erros.registrar("ERRO", "Arquivo 'Enxertado' não possui a ultima linha com registro 30")
            return False

#        log('Movendo arquivos processados ... Aguarde ...')
        # if os.path.isdir( os.path.join(dir_processados, sub_dir)) :
        #     shutil.rmtree(os.path.join(dir_processados, sub_dir))
#        if not os.path.isdir( os.path.join( dir_processados, sub_dir )):
#            os.makedirs( os.path.join( dir_processados, sub_dir ))
#        shutil.move(dnenx, os.path.join( dir_processados, sub_dir ) )
    else :
        if os.path.isfile(dnenx) :
            os.remove(dnenx)
        relatorio_erros.registrar( 'ERRO', 'Arquivos com periodos de dados diferentes.' )
        log('ERRO - Arquivos com periodos de dados diferentes ... Verifique !!')
        return False
    
    relatorio_erros.close()
    return True


def organizaArquivos(dir_trabalho):
    log("Organizando arquivos .... Aguarde ...")
    dir_regerados = os.path.join(dir_trabalho, 'REGERADOS')
    dir_protocolados = os.path.join(dir_trabalho, 'PROTOCOLADOS')
    dir_enxertados = os.path.join(dir_trabalho, 'ENXERTADOS')

    if not os.path.isdir(dir_regerados) :
        os.makedirs(dir_regerados)
    if not os.path.isdir(dir_protocolados) :
        os.makedirs(dir_protocolados)
    if not os.path.isdir(dir_enxertados) :
        os.makedirs(dir_enxertados)

    for item in os.listdir(dir_trabalho) :
        path = os.path.join(dir_trabalho, item)
        if os.path.isfile(path) :
            log('Verificando arquivo ..: %s'%(item))
            if item.startswith('GIA_') and item.__contains__('_REG_') :
                log(' - Movendo arquivo para diretorio de REGERADOS.')
                shutil.move( path, dir_regerados )
            else :
                log(' - ### Erro ... Arquivo com nome fora do padrão ... Ignorando!')

    ### Organiza Regerados ...
    ### Organiza Protocolados ...
    lst_diretorios = [ dir_regerados, dir_protocolados ]
    for dir_organizar in lst_diretorios :
        log('Organizando diretorio ..: %s'%(dir_organizar))
        for item in os.listdir(dir_organizar) :
            path = os.path.join(dir_organizar, item)
            if os.path.isfile(path) :
                log('Verificando arquivo ..: %s'%(item))
                fd = open(path, 'r', encoding = encodingDoArquivo(path))
                mes_ano = ''
                hdr = ''
                for lin in fd.readlines() :
                    if lin.startswith('05') :
                        hdr = lin
                        break
                hdr = quebraRegistro(hdr)
                mes_ano = "%s_%s"%(hdr[5][-2:], hdr[5][:4])
                dir_final = os.path.join( dir_organizar, mes_ano )
                if not os.path.isdir(dir_final) :
                    os.makedirs(dir_final)
                log(' - Diretorio destino ..: %s'%(dir_final))
                shutil.move( path, dir_final )


    log('Arquivos organizados !')
    
    return True

def ies_existentes_gia(mascara,diretorio):
    global ret
    
    qdade = 0
    ies = []
    directory = Path(diretorio)
    files = directory.glob(mascara)
    sorted_files = sorted(files, reverse=False)
    if sorted_files:
        log("Arquivos encontrados: ")
        for f in sorted_files:
            qdade = qdade + 1
            ie = str(f).split("_")[3]
            log("   ",qdade, " => ", f)
            try:
                ies.index(str(f).split("_")[2])
            except:
                ies.append(str(f).split("_")[2])
                continue
            
    else: 
        log('ERRO : Arquivo %s nao esta na pasta %s'%(mascara,diretorio))
        log("")
        ret=99
        return("")
    log(" ")
    return(ies)

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
        log('ERRO : Arquivo %s nao esta na pasta %s'%(mascara,diretorio))
        log("")
    return(nomearq)

def validauf(uf):
    return(True if (uf.upper() in ('AC','AL','AM','AP','BA','CE','DF','ES','GO','MA','MG','MS','MT','PA','PB','PE','PI','PR','RJ','RN','RO','RR','RS','SC','SE','SP','TO')) else False)
          
def dtf():
    return (datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'))

def valida_ie(ie):
    ie = re.sub('[^0-9]','',ie)
    return( "#" if ( (ie == "") or (ie == "''") or (ie == '""') or (int("0"+ie) == 0)) else ie )

def valida_mesano(mesano):
    return (mesano if (len(mesano)==6  
        and int(mesano[:2])>0 
        and int(mesano[:2])<13
        and int(mesano[2:])<=datetime.datetime.now().year
        and int(mesano[2:])>(datetime.datetime.now().year)-50) else "#" )
    
def valida_uf(uf):
    return(uf.upper() if(validauf(uf)) else "#")

def validaano(ano):
    if ano.isdigit():
        return(True if (len(ano) == 4 and int(ano) > 2000 and int(ano) <= (datetime.datetime.now().year )) else False)
    return(False)

def processo() :
    global datref
    global iei
    global ufi

#### Recebe, verifica e formata os argumentos de entrada.
    ret = 0
    ufi = ""
    mesanoi = ""
    mesi = ""
    anoi = "" 
    iei = ""

    if (len(sys.argv) == 4):
        ufi = valida_uf(str(sys.argv[1]))
        mesanoi = valida_mesano(str(sys.argv[2]))
        iei = valida_ie(str(sys.argv[3]))
    elif (len(sys.argv) == 3):
        ufi = valida_uf(str(sys.argv[1]))
        mesanoi = valida_mesano(str(sys.argv[2]))
    elif (len(sys.argv) == 2):
        ufi = valida_uf(str(sys.argv[1]))
    else:
        ret = 99

    if ( ufi == "#" or mesanoi == "#" or iei == "#" ):
        ret = 99
    
    if ( mesanoi == "" ):
        mesanoi = "*"
    if ( iei == "" ):
        iei = "*"
    if ( ufi == "" ):
        ufi = "*"
    datref = "01"+mesanoi
    
    
    if ( ret != 0):
        log("-" * 100)
        log("#### ")
        log('#### ERRO - Erro nos parametros do script.')
        log("#### ")
        log('#### Exemplo de como deve ser :')
        log('####      %s <UF> [MMAAAA] [IE] '%(sys.argv[0]))
        log("#### ")
        log('#### Onde')
        log('####      <UF>     = Obrigtório. Estado. Ex: SP')
        log('####      [MMAAAA] = Opcional. Mês e ano. Ex: Para junho de 2020 informe 062020')
        log('####      [IE]     = Opcional. Inscrição Estadual. Caso informado, o MMAAAA também deve ser informado.')
        log("#### ")
        log('#### Portanto, se o estado = SP, o mes = 06 e o ano = 2020, e IE = 108383949112 o comando correto deve ser :')  
        log('####      %s SP 062020 108383949112'%(sys.argv[0]))  
        log('#### Outros exemplos válidos:')  
        log('####      %s SP '%(sys.argv[0]))         
        log('####      %s SP 062020 '%(sys.argv[0]))         
        log("#### ")
        log('#### ')
        log("-" * 100)
        log("")
        log("Retorno = 99") 
        return(False)
      

    log('Carregando arquivo de configuracoes ...')
    
    #dir_base = disco + SD + 'arquivos' + SD + 'GIA' + ufi + SD
    dir_dados        = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'GIA') # [PTITES-1634] # configuracoes.diretorio_arquivos.replace("<<uf>>",ufi)+SD
    dir_protocolados = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'GIA', 'PROTOCOLADOS', ufi) # [PTITES-1634] # os.path.join(dir_dados.split('enxerto_gia')[-1], 'PROTOCOLADOS')
    dir_regerados    = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'GIA', 'REGERADOS', ufi)  # [PTITES-1634] # os.path.join(dir_dados.split('enxerto_gia')[-1], 'REGERADOS')
    dir_enxertados    = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'GIA', 'ENXERTADOS', ufi)  # [PTITES-1634] #
    
    log('Novo painel dir_dados.:',dir_dados)

    ERRO = 0
    OK = 0

    if not leIndiceMunicipioDIPAM( mesanoi[-4:], mesanoi[:2], ufi ):
        log("Erro ao buscar indices de municipio DIPAM")
        return False

    if os.path.isdir(dir_protocolados) :
        log("Varrendo diretorio GIA     :", ufi, str(dir_protocolados))
        
        for dir_ano in os.listdir( dir_protocolados ) :
            if validaano(dir_ano):
                if ((dir_ano == mesanoi[2:]) or (mesanoi == "*")):
                    path_ano = os.path.join(dir_protocolados, dir_ano)
                    if os.path.isdir(path_ano) :
                        log("Varrendo sub-diretorio ano :", path_ano)
                      
                        for dir_mes in os.listdir( path_ano ) :
                            if util.validames(dir_mes):
                                if ((dir_mes == mesanoi[0:2]) or (mesanoi == "*")):
                                    path_final = os.path.join(path_ano, dir_mes)
                                    if os.path.isdir(path_final) :
                                        log("Varrendo sub-diretorio mes :", path_final)
                                        mascara = "GIA*_*_*_PROT_V*.txt"
                                        listadeies = ies_existentes_gia(mascara,path_final)

                                        for iee in listadeies:
                                            if ((iee == iei) or (iei == "*")):
                                                log("Inicio do processamento para a IE ", iee)
                                                mascara_protocolado = "GIA"+ ufi +"_"+ str(dir_mes) + str(dir_ano) + "_" + iee + "_PROT_V*" + ".txt"  
                                                nome_protocolado = nome_arquivo(mascara_protocolado,path_final)
                                                if (os.path.isfile(nome_protocolado)):
                                                    log("Nome arquivo protocolado selecionado = ", nome_protocolado)
                                                    item = str(nome_protocolado).split('/')[6] 
                                                    path_regerados = os.path.join(dir_regerados, dir_ano ,dir_mes) # PTITES-1634 #path_final.replace("PROTOCOLADOS", "REGERADOS") 
                                                    mascara_regerado = "GIA"+ ufi +"_"+ str(dir_mes) + str(dir_ano) + "_" + iee + "_REG_V*" + ".txt"  
                                                    nome_regerado = nome_arquivo(mascara_regerado,path_regerados)
                                                    if (os.path.isfile(nome_regerado)):
                                                        log("Nome arquivo regerado selecionado    = ", nome_regerado)
                                                        verreg = '{:03d}'.format(int((str(nome_regerado).split(".")[0]).split("_")[4][1:]))
                                                        denx = os.path.join(dir_enxertados, str(dir_ano),str(dir_mes)) # PTITES-1634 # 
                                                        dhr = str(nome_regerado).split('_')[-1]
                                                        nenx  = "GIA"+ ufi +"_"+ str(dir_mes) + str(dir_ano) + "_" + iee + "_ENX_V" + verreg + ".txt"  
                                                        nrerr = "GIA"+ ufi +"_"+ str(dir_mes) + str(dir_ano) + "_" + iee + "_ERR_V" + verreg + ".txt"
                                                        if processaDiretorio(nome_protocolado, nome_regerado, nenx, nrerr, denx):
                                                            OK = OK + 1
                                                        else:
                                                            ERRO = ERRO + 1            
                                                    else:                
                                                        ERRO = ERRO + 1   
                                                else:                
                                                    ERRO = ERRO + 1    
            log("")                                                             
            log("Proximo ...")                                                     
        log("Fim dos dados a serem processados.")
        
    log("")
    if (OK > 0):
        log("="* 67)
        log("### SUCESSO : %s arquivo(s) processado(s)."%(OK))
        log("="* 67)
    if (ERRO > 0):
        log("ERRO : %s arquivo(s) com erro(s) não processado(s)."%(ERRO))
                                                             
    return
 
if __name__ == "__main__":
    log("Iniciando o processamento ...")
    
    comum.carregaConfiguracoes(configuracoes)
    banco=sql.geraCnxBD(configuracoes)
    
    if not criarDicionarioLayouts() :
        log("ERRO")

    processo()
    log("")
    log("FIM do processamento ...")
