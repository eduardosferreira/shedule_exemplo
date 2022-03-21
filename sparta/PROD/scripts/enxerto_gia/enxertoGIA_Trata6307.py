#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: GF
  MODULO ...: 
  SCRIPT ...: enxertoGIA.py
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
    
    * 14/04/2020 - Welber Pena de Sousa - Kyros Tecnologia 
        ### ALT001
            - Se o arquivo regerado possuir o registro tipo 14 para UF 26 no CFOP 6307:
                            - Se possuir algum outro registro 14 para este CFOP:
                                    retira o registro 14 e atualiza os valores no registro tipo 10 subtraindo os valores do tipo 14
                            - Senão
                                    Retira o registro 14 e o registro tipo 10

----------------------------------------------------------------------------------------------
"""

import os
import sys
import cx_Oracle
import unicodedata
import fnmatch
import shutil
import datetime

sys.path.append('./API_Python')

import databases
import config
import API_Python.logger as logger

name_script = os.path.basename(__file__).split('.')[0]

log = logger.startLog(name_script, debug = False )
__builtins__.log = log

Config = config.Config(name_script)

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

def encodingDoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r', encoding='iso-8859-1')
        t = fd.read()
        fd.close()
    except :
        return 'utf-8'

    return 'iso-8859-1'

def criarDicionarioLayouts() :
    log.info('Gerando dicionario de Layouts.')
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
    log.info('Dicionario de Layouts criado !')
    return True

def leIndiceMunicipioDIPAM() :
    log.info('Gerando dicionario de Indices DIPAM')
    try :
        obj_arq_DIPAM_leitura = open( './IndiceMunicipioDIPAM.csv', 'r' )
        for lin in obj_arq_DIPAM_leitura :
            itens = lin.split(';')
            if itens[0].isdigit() :
                dic_indicesDIPAM[int(itens[0])] = [itens[1], itens[2].replace('\r','').replace('\n','').replace(',','.').strip() ]
        obj_arq_DIPAM_leitura.close()
        log.info('Dicionario de Indices DIPAM criado !')
    except :
        log.erro('Impossivel ler o arquivo : IndiceMunicipioDIPAM.csv')
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

def processaDiretorio(dir_base, sub_dir) :

    dir_a_processar = os.path.join(dir_base, 'a_processar')
    dir_processados = os.path.join(dir_base, 'processados')
    if not os.path.isdir(dir_processados) :
        os.makedirs(dir_processados)
    
    path_trabalho = os.path.join(dir_a_processar, sub_dir)
    path_protocolados = os.path.join(path_trabalho, 'protocolados') 
    path_regerados = os.path.join(path_trabalho, 'regerados') 

    arq = os.listdir(path_protocolados)[0]
    if arq.upper().__contains__('REGERADO') :
        log.info('Arquivo do tipo REGERADO ... Ignorando !')
        return True
    path_arq = os.path.join(path_protocolados, arq)

    arq_new = 'ENXERTADO_' + arq
    path_arq_new = os.path.join(path_trabalho, arq_new )

    encoding = encodingDoArquivo(path_arq) 
    obj_arq_leitura = open(path_arq, 'r', encoding = encoding )
    obj_arq_regerado = None
    obj_arq_enxertado = open(path_arq_new, 'w', encoding = 'iso-8859-1')


    log.info('Arquivo a ser gerado ..: %s'%(arq_new))

    x = 0
    lin = obj_arq_leitura.readline() ### Le do protocolado .
    novas_linhas = []
    valorEntradas = 0
    valorSaidas = 0
    while lin :
    # for lin in obj_arq_leitura :
        
        if lin.startswith('05') :
            log.info('Encontrado novo documento .')
            registro = quebraRegistro(lin)            
            cabecalho_corrente = registro[:]
            linhas_a_escrever = []
            ### Busca o documento '05' referente no arquivo regerado.
            if obj_arq_regerado :
                obj_arq_regerado.close()
                del obj_arq_regerado
            
            #### Busca o arquivo REGERADO na pasta de arquivos regerados.
            
            arq_regerado = os.listdir(path_regerados)[0]
            path_arq_regerado = os.path.join(path_regerados, arq_regerado )
            obj_arq_regerado = open(path_arq_regerado, 'r', encoding=encodingDoArquivo(path_arq_regerado))
            achou = 0
            registro_regerado = [7]
            while achou < 1 :
                l = obj_arq_regerado.readline()
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
                retirar14 = False
                q14 = 0
                
                ### ALT001 - Variaveis utilizadas na alteracao .
                ValorContabil14 = 0
                BaseCalculo14 = 0
                
                Imposto14 = 0
                Outras14 = 0
                ImpostoRetidoST14 = 0
                # ImpRetSubstitutoST = 0   
                # ImpRetSubstituido = 0
                # Para referência > 200804
                #     Se (CFOP <> 1.360 e (1.401 a 1.449) e (1.651 a 1.699) e
                #     1.9xx e (2.401 a 2.449) e (2.651 a 2.699) e 2.9xx e 5.360 e
                #     (5.401 a 5.449) e (5.651 a 5.699) e 5.9xx e 6.360 e (6.401 a
                #     6.449) e (6.651 a 6. 6.99) e 6.9xx então os campos
                #     ImpRetSubstitutoST e ImpRetSubstituído devem ser
                #     preenchidos com ZEROS
                ##################################################

                soma = False
                linha10 = []
                while not fim_documento :
                    prox_lin_regerado = obj_arq_regerado.readline()
                    if not prox_lin_regerado.startswith('05') :
                        if prox_lin_regerado :
                            #########################################################################
                            #### Alteracao : Welber Pena - 04/12/2019

                            ### ALT001 - Inico da alteracao #########################################################################
                            if prox_lin_regerado.startswith('10') or linha10 :
                                if prox_lin_regerado.startswith('10') and linha10 :
                                    if len(linha10) > 1 :
                                        r10 = quebraRegistro(linha10[0])

                                        ### Atualiza o campo q14 na linha r10
                                        idx_q14 = dic_campos['10']['Q14']
                                        tam_q14 = dic_registros['10'][idx_q14][1]
                                        r10[idx_q14] = str(q14).rjust(tam_q14, '0')
                                        
                                        ### Atualiza o campo ValorContabil na linha r10
                                        idx_ValorContabil = dic_campos['10']['ValorContabil']
                                        tam_ValorContabil = dic_registros['10'][idx_ValorContabil][1]
                                        r10[idx_ValorContabil] = str(ValorContabil14).rjust(tam_ValorContabil, '0')
                                        
                                        ### Atualiza o campo BaseCalculo na linha r10
                                        idx_BaseCalculo = dic_campos['10']['BaseCalculo']
                                        tam_BaseCalculo = dic_registros['10'][idx_BaseCalculo][1]
                                        r10[idx_BaseCalculo] = str(BaseCalculo14).rjust(tam_BaseCalculo, '0')
                                        
                                        ### Atualiza o campo Imposto na linha r10
                                        idx_Imposto = dic_campos['10']['Imposto']
                                        tam_Imposto = dic_registros['10'][idx_Imposto][1]
                                        r10[idx_Imposto] = str(Imposto14).rjust(tam_Imposto, '0')

                                        ### Atualiza o campo Outras na linha r10
                                        idx_Outras = dic_campos['10']['Outras']
                                        tam_Outras = dic_registros['10'][idx_Outras][1]
                                        r10[idx_Outras] = str(Outras14).rjust(tam_Outras, '0')

                                        ### Atualiza o campo ImpostoRetidoST na linha r10
                                        idx_ImpostoRetidoST = dic_campos['10']['ImpostoRetidoST']
                                        tam_ImpostoRetidoST = dic_registros['10'][idx_ImpostoRetidoST][1]
                                        r10[idx_ImpostoRetidoST] = str(ImpostoRetidoST14).rjust(tam_ImpostoRetidoST, '0')

                                        # ### Atualiza o campo ImpRetSubstitutoST na linha r10
                                        # idx_ImpRetSubstitutoST = dic_campos['10']['ImpRetSubstitutoST']
                                        # tam_ImpRetSubstitutoST = dic_registros['10'][idx_ImpRetSubstitutoST][1]
                                        # r10[idx_ImpRetSubstitutoST] = str(ImpRetSubstitutoST14).rjust(tam_ImpRetSubstitutoST, '0')

                                        # ### Atualiza o campo ImpRetSubstituido na linha r10
                                        # idx_ImpRetSubstituido = dic_campos['10']['ImpRetSubstituido']
                                        # tam_ImpRetSubstituido = dic_registros['10'][idx_ImpRetSubstituido][1]
                                        # r10[idx_ImpRetSubstituido] = str(ImpRetSubstituido14).rjust(tam_ImpRetSubstituido, '0')

                                        ### Atualiza a linha r10 inicial com os valores calculados acima.
                                        linha10[0] = "".join(x for x in r10)

                                        for l_tmp in linha10 :
                                            novas_linhas.append(l_tmp)

                                    linha10 = []

                                if not linha10 :
                                    r10 = quebraRegistro(prox_lin_regerado)
                                    idx_cfop = dic_campos['10']['CFOP']
                                    cfop = r10[idx_cfop][:4]
                                    soma = True if r10[idx_cfop].startswith('6') else False
                                    q14 = int(r10[dic_campos['10']['Q14']])

                                    if int(cfop) == 6307 :
                                        linha10 = [prox_lin_regerado]
                                        q14 = 0
                                        ValorContabil14 = 0
                                        BaseCalculo14 = 0

                                        Imposto14 = 0
                                        Outras14 = 0
                                        ImpostoRetidoST14 = 0
                                        # ImpRetSubstitutoST = 0
                                        # ImpRetSubstituido = 0

                                        retirar14 = False
                                else :
                                    if prox_lin_regerado[:2] == '14' :
                                        r14 = quebraRegistro(prox_lin_regerado)
                                        uf = r14[dic_campos['14']['UF']]
                                        if uf in ('26', 26) :
                                            retirar14 = True
                                        else :
                                            linha10.append(prox_lin_regerado)
                                            q14 += 1
                                            ValorContabil14 += int(r14[dic_campos['14']['Valor_Contabil_1']])
                                            BaseCalculo14 += int(r14[dic_campos['14']['BaseCalculo_1']])

                                            Imposto14 += int(r14[dic_campos['14']['Imposto']])
                                            Outras14 += int(r14[dic_campos['14']['Outras']])
                                            ImpostoRetidoST14 += int(r14[dic_campos['14']['ICMSCobradoST']])

                                            # ImpRetSubstitutoST += int(r14[dic_campos['14']['']])
                                            # ImpRetSubstituido += int(r14[dic_campos['14']['Outras']])

                                            if soma :
                                                ValorContabil14 += int(r14[dic_campos['14']['Valor_Contabil_2']])
                                                BaseCalculo14 += int(r14[dic_campos['14']['BaseCalculo_2']])
                                            
                                            retirar14 = False
                                    if prox_lin_regerado[:2] == '18' and not retirar14 :
                                        linha10.append(prox_lin_regerado)
                                                                        
                                if linha10 :
                                    continue
                            ### ALT001 - FIM da alteracao #########################################################################
                        
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
                    if not fim_documento and linha10 :
                        if len(linha10) > 1 :
                            r10 = quebraRegistro(linha10[0])

                            ### Atualiza o campo q14 na linha r10
                            idx_q14 = dic_campos['10']['Q14']
                            tam_q14 = dic_registros['10'][idx_q14][1]
                            r10[idx_q14] = str(q14).rjust(tam_q14, '0')
                            
                            ### Atualiza o campo ValorContabil na linha r10
                            idx_ValorContabil = dic_campos['10']['ValorContabil']
                            tam_ValorContabil = dic_registros['10'][idx_ValorContabil][1]
                            r10[idx_ValorContabil] = str(ValorContabil14).rjust(tam_ValorContabil, '0')
                            
                            ### Atualiza o campo BaseCalculo na linha r10
                            idx_BaseCalculo = dic_campos['10']['BaseCalculo']
                            tam_BaseCalculo = dic_registros['10'][idx_BaseCalculo][1]
                            r10[idx_BaseCalculo] = str(BaseCalculo14).rjust(tam_BaseCalculo, '0')
                            
                            ### Atualiza o campo Imposto na linha r10
                            idx_Imposto = dic_campos['10']['Imposto']
                            tam_Imposto = dic_registros['10'][idx_Imposto][1]
                            r10[idx_Imposto] = str(Imposto14).rjust(tam_Imposto, '0')

                            ### Atualiza o campo Outras na linha r10
                            idx_Outras = dic_campos['10']['Outras']
                            tam_Outras = dic_registros['10'][idx_Outras][1]
                            r10[idx_Outras] = str(Outras14).rjust(tam_Outras, '0')

                            ### Atualiza o campo ImpostoRetidoST na linha r10
                            idx_ImpostoRetidoST = dic_campos['10']['ImpostoRetidoST']
                            tam_ImpostoRetidoST = dic_registros['10'][idx_ImpostoRetidoST][1]
                            r10[idx_ImpostoRetidoST] = str(ImpostoRetidoST14).rjust(tam_ImpostoRetidoST, '0')

                            # ### Atualiza o campo ImpRetSubstitutoST na linha r10
                            # idx_ImpRetSubstitutoST = dic_campos['10']['ImpRetSubstitutoST']
                            # tam_ImpRetSubstitutoST = dic_registros['10'][idx_ImpRetSubstitutoST][1]
                            # r10[idx_ImpRetSubstitutoST] = str(ImpRetSubstitutoST14).rjust(tam_ImpRetSubstitutoST, '0')

                            # ### Atualiza o campo ImpRetSubstituido na linha r10
                            # idx_ImpRetSubstituido = dic_campos['10']['ImpRetSubstituido']
                            # tam_ImpRetSubstituido = dic_registros['10'][idx_ImpRetSubstituido][1]
                            # r10[idx_ImpRetSubstituido] = str(ImpRetSubstituido14).rjust(tam_ImpRetSubstituido, '0')

                            ### Atualiza a linha r10 inicial com os valores calculados acima.
                            linha10[0] = "".join(x for x in r10)

                            for l_tmp in linha10 :
                                novas_linhas.append(l_tmp)

            #### Encontrou as linhas referentes no arquivo regerado.

            #### Pega todas as linhas referentes ao documento.
            log.info('Buscando todas as linhas referentes ao documento.')
            linhas_protocolado = []
            fim_documento = False
            while not fim_documento :
                lin = obj_arq_leitura.readline()
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
                                valorEntradas += int(r10[idx_ValorContabil].lstrip('0'))
                            ####
                    else :
                        fim_documento = True
                else :
                    fim_documento = True
            #########################################################################
            #### Alteracao : Welber Pena - 04/12/2019
            log.info('Valor Entradas : %s'%( valorEntradas ))
            log.info('Valor Saidas   : %s'%( valorSaidas ))
            fator = (valorSaidas - valorEntradas ) /100.0
            log.info('Valor Fator    : %s'%( fator ))
            ##### 
        elif lin.startswith('01') :
            log.info('Gerando HEADER do arquivo enxertado.')
            obj_arq_enxertado.write(lin.replace('\n', '\r\n'))
            lin = obj_arq_leitura.readline()
        if novas_linhas and linhas_protocolado and not linhas_a_escrever :
            q10 = 0
            log.info('Escrevendo registros do documento.')
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
                                    log.erro('Codigo de municipio invalido ! < %s >'%(r30[idx_municipio]))
                                    return False
                                
                                indice = float( indice )
                                valor_DIPAM = (fator * indice) / 100.0
                                r30[idx_valor] = str('%016.2f'%(round(valor_DIPAM, 2))).replace('.','')
                                l = linhas_protocolado.pop(y_linhas_protocolado)
                                linhas_a_escrever.append( ''.join( x for x in r30) )
                                log.info('Recalculado valor para o municipio %s - %s - %016.2f >'%( municipio, r30[idx_municipio], round(valor_DIPAM, 2) ))
                                log.info('  - Indice : %s | Fator : %s | Valor : %s'%(indice, fator, valor_DIPAM)) 
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
                obj_arq_enxertado.write( "".join(x for x in cabecalho_corrente).replace('\n', '\r\n') )
                for l in linhas_a_escrever :
                    obj_arq_enxertado.write( l.replace('\n', '\r\n') )
                    x += 1
    # print(lin)

    log.info('Arquivo processado !')
    obj_arq_leitura.close()
    obj_arq_enxertado.close()

    log.info(" R E S U M O ".center(50,'*'))
    log.info('** Numero de linhas processadas ..: %s'%(x))
    # log.info('** Qtde de doc nao encontrados ...: %s'%(linha_doc_nao_encontrado))
    # log.info('** Qtde de registros alterados ...: %s'%(atualizados))
    # log.info('** Qtde de reg ja alterados ......: %s'%(qtd_ja_alterados))
    # log.info('** Qtde de cad_cli inseridos .....: %s'%(qtd_cli_fornec_inseridos))
    # log.info('** Qtde de complvu inseridos .....: %s'%(qtd_complvu_inseridos))
    log.info('*'*50)

    log.info('Movendo arquivos processados ... Aguarde ...')
    if os.path.isdir( os.path.join(dir_processados, sub_dir)) :
        shutil.rmtree(os.path.join(dir_processados, sub_dir))
    shutil.move(path_trabalho, dir_processados )

    return True

def processo() :
    log.info("Iniciando o processamento ...")
    log.info(' - Carregando arquivo de configuracoes ...')

    dir_base = Config.getItem('diretorio_arquivos') or '.'
    dir_base = '.'
    dir_trabalho = os.path.join(dir_base, 'a_processar')

    for sub_dir in os.listdir(dir_trabalho) :
        path_trabalho = os.path.join(dir_trabalho, sub_dir)
        print(path_trabalho)
        if os.path.isdir(path_trabalho) :
            path_protocolados = os.path.join(path_trabalho, 'protocolados') 
            path_regerados = os.path.join(path_trabalho, 'regerados') 
            if os.path.isdir(path_protocolados) and os.path.isdir(path_regerados) :
                log.info('-'*100)
                log.info('Processando diretorio ..: %s'%(path_trabalho))
                processaDiretorio(dir_base, sub_dir)


    log.info("FIM do processamento ...")
    return True

if __name__ == "__main__":
    if not criarDicionarioLayouts() :
        log.finaliza("ERRO")
    if not leIndiceMunicipioDIPAM() :
        log.finaliza("ERRO")
    if not processo() :
        log.finaliza("ERRO")
    log.finaliza("SUCESSO")


