#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: Enxerto_SPED_Novo
  CRIACAO ..: 27/10/2020
  AUTOR ....: Airton Borges da Silva Filho / KYROS Consultoria
  DESCRICAO : 
----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 16/10/2020 - Airton Borges da Silva Filho / KYROS Consultoria - Criacao do script.
    * 09/12/2020 - Airton - Incluir biblioteca re, verificar IE informado se não é aspas.
    * 08/06/2021 - Airton - Regras para Blocos 1400 e 1600
    * 16/08/2021 - Adaptação para o novo Painel de execuções - Airton Borges da Silva Filho / Kyros Tec.
    * 12/11/2021 - Corrigido o problema de não gravar os registros 0150 dos 1600 que foram para o enxertado e não existia 0150.
    * 31/01/1980 - Eduardo da Silva Ferreira / Kyros Tecnologia : eduardof@kyros.com.br
                   TAG: [PTITES-1447]
                   https://jira.telefonica.com.br/browse/PTITES-1447
                   https://wikicorp.telefonica.com.br/x/QpiGDQ
                   DV - Inclusão do Cálculo do Registro E110 no Processo de Enxerto SPED 
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
from pathlib import Path
from openpyxl import load_workbook
import shutil
global ret
global series_troca


toti = {}
totc = {}
totais = {}

toti['9900'] = 4

relatorio_erros = None

# inicio [PTITES-1447]
v_soma_vl_icms_protocolado = float(0)
v_soma_vl_icms_regerado = float(0)
v_flag = False

v_fl_processado_regerado = False
v_fl_encontrou_c190_protocolado = False
v_fl_encontrou_d696_regerado = False
# fim [PTITES-1447]

log.apagar_arq = False

class Relatorio:
    def __init__(self, nome_relatorio, dir_geracao = ''):
        self.__arquivo_relatorio = '%s_%s.csv'%( nome_relatorio.capitalize(), datetime.datetime.now().strftime('%Y%m%d') )
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
            # log('Registrar :', args)
            self.__linhas_relatorio.append([x for x in args])
        else :
            log("ERRO ao registar dados no relatorio.")
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


def contarLinhasArquivo(nome_arquivo):
    def blocks(files, size=65536):
        while True:
            b = files.read(size)
            if not b: break
            yield b
    encoding = comum.encodingDoArquivo(nome_arquivo)
#    encoding = encodingDoArquivo(nome_arquivo)
    with open(nome_arquivo, "r", encoding=encoding, errors='ignore') as f:
        return sum(bl.count("\n") for bl in blocks(f))

def formatNumero(numero): return '{:,}'.format(numero).replace(',', '.')


class Bloco_paraCopia:
    __blocos_paraCopia = None
    __linhasBloco = None
    __qtd_OutrasLinhasBlocoDestino = 0
    __qtd_LinhasIniciadas9 = 0
    __qtd_LinhasIniciadas9900 = 0
    __inserir_noFinal = True
    __letra = None
    __naoTratarBloco = False
    __ancora_chave = None
    __blocoMarcado = False

    def __init__(self, letra):
        self.__letra = letra
        self.__linhasBloco = []
        self.__blocos_paraCopia = {}
        self.inserir_noFinal = False

    @property
    def blocos_paraCopia(self): return self.__blocos_paraCopia
    @property
    def linhasBloco(self): return self.__linhasBloco
    @property
    def qtd_OutrasLinhasBlocoDestino(self): return self.__qtd_OutrasLinhasBlocoDestino
    @property
    def qtd_LinhasIniciadas9(self): return self.__qtd_LinhasIniciadas9
    @property
    def qtd_LinhasIniciadas9900(self): return self.__qtd_LinhasIniciadas9900
    @property
    def qtd_linhasCopiadas(self): return len(self.linhasBloco)
    @property
    def qtd_linhasTotaisBloco(self): return self.qtd_linhasCopiadas + self.__qtd_OutrasLinhasBlocoDestino
    @property
    def inserir_noFinal(self): return self.__inserir_noFinal
    @inserir_noFinal.setter
    def inserir_noFinal(self, valor):
        self.__inserir_noFinal = valor
        if valor:
            self.ancora_chave = self.__letra + '990'
        else:
            self.ancora_chave = self.__letra + '001'
    @property
    def naoTratarBloco(self): return self.__naoTratarBloco
    @naoTratarBloco.setter
    def naoTratarBloco(self, valor): self.__naoTratarBloco = valor
    @property
    def ancora_chave(self): return self.__ancora_chave
    @ancora_chave.setter
    def ancora_chave(self, valor): self.__ancora_chave = valor

    def inc_OutrasLinhasBlocoDestino(self):
        self.__qtd_OutrasLinhasBlocoDestino += 1

    def inc_Linhasiniciadas9(self):
        self.__qtd_LinhasIniciadas9 += 1
    
    def inc_Linhasiniciadas9900(self):
        self.__qtd_LinhasIniciadas9900 += 1

    def processarLinha_Origem(self, txt_linha):
        encontrado = False
        global series_troca
        for blc in self.blocos_paraCopia:
            if txt_linha.startswith('|' + blc + '|'):
                ############ ALTERADO Inicio
                if series_troca :
                    if blc == 'D695' :
                        if txt_linha.split('|')[3] in series_troca :
                            self.__blocoMarcado = True
                        else :
                            self.__blocoMarcado = False
                else :
                    self.__blocoMarcado = True
                if self.__blocoMarcado :
                ############ ALTERADO fim
                    self.blocos_paraCopia[blc] += 1
                    self.linhasBloco.append(txt_linha)
                    encontrado = True
                break
        return encontrado
        #         encontrado = True
        #         break
        # return encontrado
        ############ ALTERADO                            

            


class Blocos_paraCopia:
    __blocos = None
    registroAncora = None
    blocoMarcado = False

    def __init__(self, *listaBlocos):
        self.__blocos = {}
        for nomeBloco in listaBlocos:
            if not nomeBloco[0] in self.__blocos:
                blc = Bloco_paraCopia(nomeBloco[0])
                self.__blocos[nomeBloco[0]] = blc
            else:
                blc = self.__blocos[nomeBloco[0]]
            blc.blocos_paraCopia[nomeBloco] = 0

    @property
    def blocos(self): return self.__blocos

    def processarLinha_Origem(self, txt_linha):
        for blc in self.blocos:
            if self.blocos[blc].naoTratarBloco: continue
            if self.blocos[blc].processarLinha_Origem(txt_linha): break

    def processarLinha_Destino(self, txt_linha, num_linha, arq_novo):
        global series_troca
        # inicio [PTITES-1447]
        global v_soma_vl_icms_regerado
        global v_fl_processado_regerado
        global v_fl_encontrou_d696_regerado
        # fim [PTITES-1447]
        
        for letraBloco in self.blocos:
            bloc = self.blocos[letraBloco]
            # inicio [PTITES-1447]
            # if not v_fl_processado_regerado:
            #     """
            #         Implementar a lógica abaixo dentro da função processarLinha_Destino da classe no loop de faz a leitura das linhas do bloco a enxertar:
            #         Algorítimo :
            #         Se o campo 1 do registro for igual a D696:
            #             v_soma_vl_icms_regerado = v_soma_vl_icms_regerado + o campo 7 do registro
            #     """
            #     for line in bloc.linhasBloco:
            #         if line.startswith('|D696|'):
            #             v_soma_vl_icms_regerado += float(str(line.split('|')[7]).replace(',', '.'))
            #             v_fl_encontrou_d696_regerado = True
            #     v_fl_processado_regerado = True         
            # fim [PTITES-1447]
                        
            if bloc.naoTratarBloco: continue

            ## Checando se a linha pertence a um bloco que será sobrescrito
            if txt_linha.startswith('|'+letraBloco):
                for cod_bloco in bloc.blocos_paraCopia:
                    if txt_linha.startswith("|" + cod_bloco + "|"):
                        # return True
                        ############ ALTERADO
                        if series_troca :
                            if cod_bloco == 'D695' :
                                if txt_linha.split('|')[3] in series_troca :
                                    self.blocoMarcado = True
                                else :
                                    self.blocoMarcado = False
                        else :
                            self.blocoMarcado = True
                        
                        if self.blocoMarcado :
                            return self.blocoMarcado
                        ############ ALTERADO
                    

                    if not v_fl_processado_regerado:
                        """
                            Implementar a lógica abaixo dentro da função processarLinha_Destino da classe no loop de faz a leitura das linhas do bloco a enxertar:
                            Algorítimo :
                            Se o campo 1 do registro for igual a D696:
                                v_soma_vl_icms_regerado = v_soma_vl_icms_regerado + o campo 7 do registro
                        """
                        for line in bloc.linhasBloco:
                            if line.startswith('|D696|'):
                                v_soma_vl_icms_regerado += float(str(line.split('|')[7]).replace(',', '.'))
                                v_fl_encontrou_d696_regerado = True
                        v_fl_processado_regerado = True    


                # bloc.qtd_linhasTotaisBloco += 1
                # log( "LIN", txt_linha)
                if txt_linha.startswith('|' + bloc.ancora_chave + '|'): ## Exemplo de linha : |D990|<qtde de linhas>
                    if not bloc.inserir_noFinal:
                        arq_novo.write(txt_linha)
                        bloc.inc_OutrasLinhasBlocoDestino()
                        log(formatNumero(num_linha), "Inserindo linhas no começo do bloco " + letraBloco)
                    else:
                        log(formatNumero(num_linha), "Inserindo linhas no final do bloco " + letraBloco)
                    
                    log('-'*70)
                    log('AKI - Qtde :', len(bloc.linhasBloco))
                    log('    - Qtde |9990 =', bloc.qtd_LinhasIniciadas9900)
                    log('    - Qtde |9    =', bloc.qtd_LinhasIniciadas9)
                    for line in bloc.linhasBloco :
                        arq_novo.write(line)
                        if line.startswith('|9990|'):
                            bloc.inc_Linhasiniciadas9900()
                        if line.startswith('|9'):
                            bloc.inc_Linhasiniciadas9()
                    log('    - Qtde |9990 =', bloc.qtd_LinhasIniciadas9900)
                    log('    - Qtde |9    =', bloc.qtd_LinhasIniciadas9)
                    log('-'*70)
                    # log('>>>>>>>>>\n', bloc.linhasBloco, '\n<<<<<<<<<<<<')

                    # arq_novo.writelines(bloc.linhasBloco)

                    if bloc.inserir_noFinal and txt_linha.startswith("|" +letraBloco+ "990|"):
                        log('Escreveu %s990 por AKI !!!! %s'%(letraBloco, bloc.qtd_OutrasLinhasBlocoDestino ))
                        arq_novo.write("|" + letraBloco + "990|" + str(bloc.qtd_linhasTotaisBloco+1) + "|\n")
                        bloc.inc_OutrasLinhasBlocoDestino()
                        return True

                    if not bloc.inserir_noFinal: return True
                elif txt_linha.startswith("|" +letraBloco+ "990|"):
                    log('FOI ESCRITO %s990 por AKI !!!!'%(letraBloco))
                    arq_novo.write("|" + letraBloco + "990|" + str(bloc.qtd_linhasTotaisBloco+1) + "|\n")
                    bloc.inc_OutrasLinhasBlocoDestino()
                    return True
                break
            ## Checando se a linha pertence a um totalizador de bloco que será reescrito
            elif txt_linha.startswith('|9900|'+letraBloco):
                # if txt_linha.startswith('|9'):
                # bloc.inc_Linhasiniciadas9()
                # if txt_linha.startswith('|9900|'):
                # bloc.inc_Linhasiniciadas9900()

                if not bloc.inserir_noFinal and txt_linha.startswith('|9900|' + bloc.ancora_chave + '|'):
                    arq_novo.write(txt_linha)
                    if txt_linha.startswith('|9'):
                        bloc.inc_Linhasiniciadas9()
                    if txt_linha.startswith('|9900|'):
                        bloc.inc_Linhasiniciadas9900()

                    bloc.inc_OutrasLinhasBlocoDestino()

                    for chave_reg in bloc.blocos_paraCopia:
                        quantidade = bloc.blocos_paraCopia[chave_reg]
                        if quantidade != 0:
                            arq_novo.write("|9900|" + chave_reg + "|" + str(quantidade) + "|\n")
                            bloc.inc_Linhasiniciadas9()
                            bloc.inc_Linhasiniciadas9900()
                            bloc.inc_OutrasLinhasBlocoDestino()
                    return True

                for cod_bloco in bloc.blocos_paraCopia:
                    if txt_linha.startswith("|9900|" + cod_bloco+ "|"):
                        return True

                ## Inserindo linhas de totalizadores
                if txt_linha.startswith('|9900|' + letraBloco + '990|'):
                    if bloc.inserir_noFinal:
                        log(formatNumero(num_linha), "Inserindo linhas dos totalizadores do bloco "+letraBloco)
                        for chave_reg in bloc.blocos_paraCopia:
                            quantidade = bloc.blocos_paraCopia[chave_reg]
                            if quantidade != 0:
                                arq_novo.write("|9900|" + chave_reg + "|" + str(quantidade) + "|\n")
                                bloc.inc_Linhasiniciadas9()
                                bloc.inc_Linhasiniciadas9900()
                                bloc.inc_OutrasLinhasBlocoDestino()
                    arq_novo.write('|9900|' + letraBloco + '990|1|\n')
                    bloc.inc_Linhasiniciadas9()
                    bloc.inc_Linhasiniciadas9900()
                    bloc.inc_OutrasLinhasBlocoDestino()
                    return True
                break

        return False

    @property
    def qtd_linhas_geradas_blocos(self):
        qtd = 0
        for blc in self.blocos: qtd += self.blocos[blc].qtd_linhasTotaisBloco
        
        return qtd


# =Até 20211116 era este:============================================================================
# def encodingDoArquivo(path_arq) :
#     global ret
#     
#     try :
#         fd = open(path_arq, 'r', encoding='iso-8859-1')
#         fd.read()
#         fd.close()
#     except :
#         return 'utf-8'
# 
#     return 'iso-8859-1'
# 
# =============================================================================
# =============================================================================
# def encodingDoArquivo(path_arq) :
#     global ret
#      
#     try :
#         fd = open(path_arq, 'r', encoding='utf-8')
#         fd.read()
#         fd.close()
#     except :
#         print("==========>>>>>>>>>>          RECONHECEU O ARQUIVO ",path_arq, " COMO : iso-8859-1" )
#         return 'iso-8859-1'
# 
#     print("==========>>>>>>>>>>          RECONHECEU O ARQUIVO ",path_arq, " COMO : utf-8" )
#     return 'utf-8'
# # 
# =============================================================================




def retornaUFArquivo(path) :
    try :
        fd = open(path,'r', encoding=comum.encodingDoArquivo(path))
#        fd = open(path,'r', encoding=encodingDoArquivo(path))
        lin = fd.readline()
    except :
        fd = open(path,'r', encoding=comum.encodingDoArquivo(path))
#        fd = open(path,'r', encoding=encodingDoArquivo(path))
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
        print ("# Arquivos encontrados: ")
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

def processar(ufi,mesanoi,mesi,anoi,iei):
    global ret
    
    nome_protocolado=""
    nome_regerado=""
    nome_enxertado=""
    dir_base = SD + 'arquivos' + SD + 'SPED_FISCAL'
    dir_protocolados = os.path.join(os.path.dirname(configuracoes.dir_entrada), 'SPED_FISCAL', 'PROTOCOLADOS', ufi, anoi, mesi) # [PTITES-1637] #
    dir_regerados = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'REGERADOS', ufi, anoi, mesi) # [PTITES-1637] #
    dir_enxertados = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'ENXERTADOS',ufi,anoi,mesi) # [PTITES-1637] #
     
    
    dir_dev = os.getcwd()
    
    dir_enxertados = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'SPED_FISCAL', 'ENXERTADOS',ufi,anoi,mesi) # [PTITES-1637] #
  
    # [PTITES-1637] #if (configuracoes.ambiente == 'DEV'):
    # [PTITES-1637] #    dir_enxertados = dir_dev+dir_enxertados
  

    
    
    
    
    
    mascara_regeradoi = "SPED_"+mesanoi+"_"+ufi+"_"+iei+"_REG*.txt"
    listadeies = ies_existentes(mascara_regeradoi,dir_regerados)
    
    #log(" = ", )
    #log(" = ", )
    #log(" = ", )
    #log("listadeies = ", listadeies)
    #log(" = ", )
    
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
            log("ERRO:    Não foi processado o ENXERTO para a dupla de arquivos:")
            log("####          Arquivo regerado    = ", nome_regerado)
            log("####          Arquivo protocolado = ", nome_protocolado)
            log("-"*100)
            ret=99
        else:
                
            ### prepara saida ENXERTADO

            if (str(nome_regerado).count("_") == 6):
                versao_enxertado = "_"+(str(nome_regerado).split(".")[0]).split("_")[6]
            else:
                versao_enxertado = ""
                
                
               
                
               
                
               
            # [PTITES-1637] #if (configuracoes.ambiente == 'DEV'):
            # [PTITES-1637] #    dir_enxertados = dir_dev+dir_enxertados   
                
                
                                
                
               
                
               
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
            
            
#            input("Vai realizar o enxerto inicial....")
            
            
            if processaDiretorio(nome_protocolado, nome_regerado, nome_enxertado, dir_protocolados, dir_regerados, dir_enxertados) == False :
                ret = 99
                if log.apagar_arq :
                    os.system('rm %s'%( nome_enxertado ))
                
        log("-"*100)
        log("#")
        log("# FIM do processamento para a IE ", iee)
        log("#")
        log("-"*100)
        log("#")
        log("#")
        log("#")

    return(ret,nome_enxertado,nome_regerado,nome_protocolado)

def parametros():
    global ret
    ufi = "SP"
    mesanoi = ""
    iei = "*"
    mesi = ""
    anoi = "" 
    ret = 0
    series_troca = None
    
#### Recebe, verifica e formata os argumentos de entrada.  
    enx1400 = 'S'
    if (len(sys.argv) == 6
        and validauf(ufi)
        and len(sys.argv[1])==6  
        and int(sys.argv[1][:2])>0 
        and int(sys.argv[1][:2])<13
        and int(sys.argv[1][2:])<=datetime.datetime.now().year
        and int(sys.argv[1][2:])>2014
        and (sys.argv[2].upper() in ("S", "'S'", '"S"',"N", "'N'", '"N"'))
        ):
     
        mesanoi = sys.argv[1].upper()
        
        #### Parametro series -------
        if sys.argv[5] :
            series_troca = [ x.replace(' ','') for x in sys.argv[5].split(',') ]


        if len(sys.argv) == 5:
            iei=sys.argv[4].upper()
            iei = re.sub('[^0-9]','',iei)
       
            if ( (iei == "") or (iei == "''") or (iei == '""') or (int("0"+iei) == 0)):
                iei = "*"
            
    else :
        log("-" * 100)
        log("#### ")
        log('ERRO - Erro nos parametros do script.')
        log("#### ")
        log('#### Exemplo de como deve ser :')
        log('####      %s <MMAAAA> <S/N> <s/n> <IE> <SERIES> '%(sys.argv[0] if sys.argv[0][0] == '.' else '.' + SD + sys.argv[0] ))
        log("#### ")
        log('#### Onde')
        # log('####      <UF> = estado. Ex: SP')
        log('####      <MMAAAA> = mês e ano. Ex: Para junho de 2020 informe 062020.')
        log('####      <S/N>    = Deve ser informado S para enxertar o bloco 1400 ou N para que não tenha o bloco 1400.')
        log('####      <s/n>    = Deve ser informado S para enxertar o bloco 1600 regerado independente da data de origem ')
        log('####                     ou N para que considere as regras de data até 072017 e registros 0150.')
        log('####      <IE>     = Inscição Estadual. É opcional, pode ou não ser informado.')
        log('####                     caso não informado, será processado para todas IEs do estado <UF> informado.')
        log('####      <SERIES> = Series a serem substituidas nos blocos D695, D696, D697 .')
        log('####                     caso não informado, todas as series serão substituidas.')
        log("#### ")
        # log('#### Portanto, se o estado = SP, o mes = 06 e o ano = 2020, e deseja todas IEs,  o comando correto deve ser :')  
        log('#### Portanto, se o mes = 06 e o ano = 2020, deseja enxertar o bloco 1400 , bloco 1600 e IE = 108383949112,  o comando correto deve ser :')  
        log('####      %s 062020 S S 108383949112 "1, C, UT"'%(sys.argv[0] if sys.argv[0][0] == '.' else '.' + SD + sys.argv[0]))  
        log("#### ")
        log("-" * 100)
        log("")
        log("Retorno = 99") 
        ret = 99

        return(False,False,False,False,False,False)

    mesanoi = sys.argv[1].upper()
    mesi    = sys.argv[1][:2].upper()
    anoi    = sys.argv[1][2:].upper()
    enx1400 = sys.argv[2].upper()
    enx1600 = sys.argv[3].upper()
    
    return(ufi,mesanoi,mesi,anoi,iei,enx1400,enx1600,series_troca)

def processaDiretorio(nome_protocolado, nome_regerado, nome_enxertado, path_protocolados, path_regerados, path_enxertados) :
    pass # inicio [PTITES-1447]
    #substituir a posicao da string    
    def fnc_replace(p_ds_dado_linha, p_ds_novo_dado_posicao, p_nr_posicao_dado=0, p_cc_quebra_dado="|"):
        try:
            v_ob_dado_linha = str(p_ds_dado_linha).split(p_cc_quebra_dado)
            v_ob_dado_linha[p_nr_posicao_dado] = p_ds_novo_dado_posicao
            v_ds_dado_linha = str(p_cc_quebra_dado).join(v_ob_dado_linha)
            return v_ds_dado_linha            
        except:
            return p_ds_dado_linha    
    
    global v_soma_vl_icms_protocolado
    global v_soma_vl_icms_regerado
    global v_flag 
    
    global v_fl_processado_regerado 
    global v_fl_encontrou_c190_protocolado
    global v_fl_encontrou_d696_regerado
     
    
    v_soma_vl_icms_protocolado = float(0)
    v_soma_vl_icms_regerado = float(0)
    v_flag = False

    v_fl_processado_regerado = False
    v_fl_encontrou_c190_protocolado = False
    v_fl_encontrou_d696_regerado = False
    # fim [PTITES-1447]
    
    global relatorio_erros 
        
    path_nomeArquivoOriginal = nome_protocolado
    path_nomeArquivoBlocos = nome_regerado
    path_nomeArquivoNovoArq = nome_enxertado

    blocos_paraCopia = Blocos_paraCopia(
        'D695', 'D696', 'D697'
    )

    blocos_paraCopia.blocos['D'].inserir_noFinal = True
    
    encoding = comum.encodingDoArquivo(path_nomeArquivoBlocos)
#    encoding = encodingDoArquivo(path_nomeArquivoBlocos)
    arquivoCopia = open(path_nomeArquivoBlocos, 'r', encoding=encoding)
    reg_inicial_arqCopia = False
    for linha in arquivoCopia:
        if linha.startswith('|0000|') :
            reg_inicial_arqCopia = linha[:]
        blocos_paraCopia.processarLinha_Origem(linha)

    arquivoCopia.close()

    log("Registros selecionados no arquivo ", path_nomeArquivoBlocos)

    for bloco in blocos_paraCopia.blocos:
        log("Bloco:", bloco)
        bloco = blocos_paraCopia.blocos[bloco]
        for chave in bloco.blocos_paraCopia:
            log("  ", chave+':', formatNumero(bloco.blocos_paraCopia[chave]))

    log("Contando linhas do arquivo original...")
    qtdLinhasArqOriginal = contarLinhasArquivo(path_nomeArquivoOriginal)

    log("Abrindo arquivo original e escrevendo arquivo novo...")
    encoding = comum.encodingDoArquivo(path_nomeArquivoOriginal)
#    encoding = encodingDoArquivo(path_nomeArquivoOriginal)
    arquivoOriginal = open(path_nomeArquivoOriginal, 'r', encoding=encoding, errors='ignore')
    arquivoNovo = open(path_nomeArquivoNovoArq, 'w', encoding=encoding )

    num_linhaAtual = 0
    num_linhaOriginaisCopiadas = 0
    qtde_registros_9900 = 0 #### Conta todos os registros iniciados com |9900|
    qtde_registros_9990 = 0 #### Conta todos os registros iniciados com |9 + o ultimo registro |9999|

    for lnh in arquivoOriginal:#[PTITES-1447]
        # inicio [PTITES-1447]
        linha = lnh
        # Implementar a lógica abaixo dentro do laço do código onde faz a leitura do arquivo protocolado :
        """
        Se o campo 1 do registro for igual a C100 faça :
            Se o campo 2 for igual a 1 faça :
                v_flag = Verdadeiro
            Senão faça :
                v_flag = Falso
        """
        if linha.split('|')[1] in ('C100'):
            if linha.split('|')[2] == '1':
                v_flag = True
            else:
                v_flag = False        
        
        '''
        Se o campo 1 do registro for igual a C190 e v_flag for igual a Verdadeiro faça :
            v_soma_vl_icms_protocolado = v_soma_vl_icms_protocolado + o campo 7 do registro
        '''
        if v_flag and linha.startswith("|C190|"):
            v_soma_vl_icms_protocolado += float(str(linha.split('|')[7]).replace(',', '.'))
            v_fl_encontrou_c190_protocolado = True 
        
        # fim [PTITES-1447]
        
        num_linhaAtual += 1
        if num_linhaAtual == 1 and reg_inicial_arqCopia :
            if (linha.split('|')[4] != reg_inicial_arqCopia.split('|')[4]) or (linha.split('|')[5] != reg_inicial_arqCopia.split('|')[5]) :
                log("#"*80)
                log('ERRO - Arquivos com periodos de dados diferentes ... Verifique !!')
                log("#"*80)
                relatorio_erros.registrar( 'ERRO', 'Arquivos com periodos de dados diferentes.' )
                return False

        if num_linhaAtual % 500000 == 0:
            log(formatNumero(num_linhaAtual), "/", formatNumero(qtdLinhasArqOriginal))

        if not blocos_paraCopia.processarLinha_Destino(linha, num_linhaAtual, arquivoNovo):
            # inicio [PTITES-1447]
            """
            Se o campo 1 do registro for igual a E110 faça :
                VL_TOT_DEBITOS = v_soma_vl_icms_protocolado + v_soma_vl_icms_regerado
                altere o valor do campo 2 ( VL_TOT_DEBITOS ) substituindo o mesmo pelo valor de VL_TOT_DEBITOS
                v_soma1 = VL_TOT_DEBITOS + VL_AJ_DEBITOS (campo 03) + VL_TOT_AJ_DEBITOS (campo 04) + VL_ESTORNOS_CRED (campo 05)
                v_soma2 = VL_TOT_CREDITOS (campo 06) + VL_AJ_CREDITOS (campo 07) + VL_TOT_AJ_CREDITOS (campo 08) + VL_ESTORNOS_DEB (campo 09) + VL_SLD_CREDOR_ANT (campo 10)
                altere o valor do campo 11 (VL_SLD_APURADO) do registro para v_soma1 – v_soma2
                altere o valor do campo 13 (VL_ICMS_RECOLHER) do registro para campo 11 (VL_SLD_APURADO) – campo 12 (VL_TOT_DED)
            """
            if linha.startswith("|E110|"):
                ####### LOG ################################                
                log("Antes LNH = ", lnh )
                #####
                VL_TOT_DEBITOS = v_soma_vl_icms_protocolado + v_soma_vl_icms_regerado
                VL_AJ_DEBITOS = float(str(linha.split('|')[3]).replace(',', '.'))
                VL_TOT_AJ_DEBITOS = float(str(linha.split('|')[4]).replace(',', '.'))    
                VL_ESTORNOS_CRED = float(str(linha.split('|')[5]).replace(',', '.'))
                VL_TOT_CREDITOS = float(str(linha.split('|')[6]).replace(',', '.'))
                VL_AJ_CREDITOS = float(str(linha.split('|')[7]).replace(',', '.'))
                VL_TOT_AJ_CREDITOS = float(str(linha.split('|')[8]).replace(',', '.'))
                VL_ESTORNOS_DEB = float(str(linha.split('|')[9]).replace(',', '.'))
                VL_SLD_CREDOR_ANT = float(str(linha.split('|')[10]).replace(',', '.'))
                VL_TOT_DED = float(str(linha.split('|')[12]).replace(',', '.'))

                linha = fnc_replace(linha,str('%.2f'%(VL_TOT_DEBITOS)).replace('.',','),2)
                
                v_soma1 = VL_TOT_DEBITOS + VL_AJ_DEBITOS + VL_TOT_AJ_DEBITOS + VL_ESTORNOS_CRED
                v_soma2 = VL_TOT_CREDITOS + VL_AJ_CREDITOS + VL_TOT_AJ_CREDITOS + VL_ESTORNOS_DEB + VL_SLD_CREDOR_ANT
                
                VL_SLD_APURADO = round((v_soma1 - v_soma2),2)
                VL_ICMS_RECOLHER = round((VL_SLD_APURADO - VL_TOT_DED),2)
                if VL_ICMS_RECOLHER < 0 :
                    log('-VL_TOT_DEBITOS',str(VL_TOT_DEBITOS))
                    log('-VL_AJ_DEBITOS',str(VL_AJ_DEBITOS))
                    log('-VL_TOT_AJ_DEBITOS',str(VL_TOT_AJ_DEBITOS))
                    log('-VL_ESTORNOS_CRED',str(VL_ESTORNOS_CRED))
                    log('-VL_TOT_CREDITOS',str(VL_TOT_CREDITOS))
                    log('-VL_AJ_CREDITOS',str(VL_AJ_CREDITOS))
                    log('-VL_TOT_AJ_CREDITOS',str(VL_TOT_AJ_CREDITOS))
                    log('-VL_ESTORNOS_DEB',str(VL_ESTORNOS_DEB))
                    log('-VL_SLD_CREDOR_ANT',str(VL_SLD_CREDOR_ANT))
                    log('-v_soma_vl_icms_protocolado',str(v_soma_vl_icms_protocolado))
                    log('-v_soma_vl_icms_regerado',str(v_soma_vl_icms_regerado))
                    log('-VL_ICMS_RECOLHER',str(VL_ICMS_RECOLHER))                   
                    
                    log("ERRO Arquivo com problema, necessária análise da equipe tributária.\n  - VL_ICMS_RECOLHER Negativo.")
                    log.apagar_arq = True
                    return False
                
                # altere o valor do campo 11 (VL_SLD_APURADO) do registro para v_soma1 – v_soma2
                linha_aux = fnc_replace(linha,str(VL_SLD_APURADO).replace('.',','),11)
                linha = linha_aux 
                
                #altere o valor do campo 13 (VL_ICMS_RECOLHER) do registro para campo 11 (VL_SLD_APURADO) – campo 12 (VL_TOT_DED)
                linha_aux = fnc_replace(linha,str(VL_ICMS_RECOLHER).replace('.',','),13)
                linha = linha_aux 
                
                ### LOG
                log("Depois LNH = ", linha )
                log("v_soma_vl_icms_protocolado = ", str(v_soma_vl_icms_protocolado) )
                log("v_soma_vl_icms_regerado = ", str(v_soma_vl_icms_regerado) )
                # log("v_soma1 = ", str(v_soma1) )
                # log("v_soma2 = ", str(v_soma2) )
                # log("VL_TOT_DED = ", str(VL_TOT_DED) )
                log("VL_SLD_APURADO = ", str(VL_SLD_APURADO) )
                # log("APURADO = ", str(v_soma1 - v_soma2) )
                log("VL_ICMS_RECOLHER = ", str(VL_ICMS_RECOLHER) )
                # log("RECOLHER = ", str(v_soma1 - v_soma2) )
                v_soma_vl_icms_protocolado = float(0)
                v_soma_vl_icms_regerado = float(0)
                v_flag = False
                v_fl_processado_regerado = False
                v_fl_encontrou_c190_protocolado = False
                v_fl_encontrou_d696_regerado = False
            # fim [PTITES-1447]
            
            if linha.startswith("|9999|"):
                qtd_linhas = num_linhaOriginaisCopiadas + blocos_paraCopia.qtd_linhas_geradas_blocos + 1 #|9999|
                log("Escrevendo registro 9999: "+formatNumero(qtd_linhas))
                arquivoNovo.write("|9999|" + str(qtd_linhas) + "|\n")
                break
            else:
                linha_de_bloco = False
                for bloco in blocos_paraCopia.blocos :
                    if linha.startswith('|' + bloco) :
                        linha_de_bloco = bloco
                
                ##### Realizar nova somatoria dos registros |9900|

                if linha.startswith("|9900|9900"):
                    log(">>> ANTES DA SOMA >> Quantidade de registros |9900| =", qtde_registros_9900)
                    for bloco in blocos_paraCopia.blocos :
                        qtde_registros_9900 += blocos_paraCopia.blocos[bloco].qtd_LinhasIniciadas9900
                    qtde_registros_9900 += 1
                    qtde_registros_9990 += 1 ##### Como esse registro inicia com |9 tem que incrementar o |9990|
                    log(">> Quantidade de registros |9900| =", qtde_registros_9900 )
                    linha = "|9900|9900|" + str(qtde_registros_9900) + "|\n"
                                
                ##### Realizar nova somatoria dos registros |9*
                elif linha.startswith("|9990|"):
                    log(">>> ANTES DA SOMA >> Quantidade de registros |9 =", qtde_registros_9990 )
                    for bloco in blocos_paraCopia.blocos :
                        qtde_registros_9990 += blocos_paraCopia.blocos[bloco].qtd_LinhasIniciadas9
                    qtde_registros_9990 += 1
                    log(">> Quantidade de registros |9 =", qtde_registros_9990 )
                    linha = "|9990|" + str(qtde_registros_9990) + "|\n"
                
                if linha.startswith("|9900|"):
                    qtde_registros_9900 += 1
                if linha.startswith("|9"):
                    qtde_registros_9990 += 1 ##### Como esse registro inicia com |9 tem que incrementar o |9990|
                arquivoNovo.write(linha)

                if not linha_de_bloco :
                    num_linhaOriginaisCopiadas += 1
                else :
                    blocos_paraCopia.blocos[linha_de_bloco].inc_OutrasLinhasBlocoDestino()

    log("Fechando arquivos")

    for bloco in blocos_paraCopia.blocos :
        log("Linhas |9 =", blocos_paraCopia.blocos[bloco].qtd_LinhasIniciadas9)
        log("Linhas |9900 =", blocos_paraCopia.blocos[bloco].qtd_LinhasIniciadas9900)

    arquivoOriginal.close()
    arquivoNovo.close()

    arquivoNovo = open(path_nomeArquivoNovoArq, 'r',  encoding=encoding)
    ultimas_linhas = ['','']
    for lin in arquivoNovo.readlines() :
        ultimas_linhas[1] = ultimas_linhas[0]
        ultimas_linhas[0] = lin
    arquivoNovo.close()

    if not ultimas_linhas[0].startswith('|9999|') :
        if not ultimas_linhas[1].startswith('|9999|') :
            log("#"*80)
            log("Erro arquivo 'Enxertado' não possui a ultima linha com registro |9999|")
            log("#"*80)
            relatorio_erros.registrar( 'ERRO', 'Arquivo ENXERTADO não possui a ultima linha com registro |9999|' )
            return False
    
    return True





def verifica1400(uf1400,ie1400,mes1400,ano1400):
    ### Diretorio devido mudanças dos diretorio de RELATORIOS ( Roney - 23/11/2021 )
    # pasta_1400 = SD + 'arquivos' + SD + 'REGISTRO_1400' + SD + 'RELATORIOS' + SD + uf1400 + SD + str(ano1400) + SD + str(mes1400) + SD
    #- [PTITES-1637]
    # pasta_1400 = SD + 'portaloptrib' + SD + 'TESHUVA' + SD + 'RELATORIOS' + SD + 'SPED_FISCAL' + SD + uf1400 + SD + str(ano1400) + SD + str(mes1400) + SD + 'REGISTRO_1400' + SD
    pasta_1400 = os.path.join(os.path.dirname(configuracoes.dir_geracao_arquivos), 'REGISTRO_1400', 'RELATORIOS', uf1400, ano1400, mes1400)
    mask_1400  = ie1400+"_"+"Valores_Agregados_1400_"+ mesanoi+".xlsx"
    nome_1400  = nome_arquivo(mask_1400,pasta_1400)
    if (nome_1400 == ""):
        log("ERRO - Execute antes o script registro_1400 para este mes, ano e IE. ")
        ret = 99
    return(nome_1400)

    








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
    if contador % 500000 == 0:
        log(formatNumero(contador))
    
    return (contador)

def preparaenxerto1600(nomeenx, nomereg, nomepro, aaaamm, enx1600):
    ret = 0
    
    #log("Recebeu nomeenx = ", nomeenx )
    #log("Recebeu nomereg = ", nomereg )
    #log("Recebeu nomepro = ", nomepro )
    #log("Recebeu aaaamm  = ", aaaamm )
    #log("Recebeu enx1600 = ", enx1600 )

   
    ##### BLOCO 1600
    ##### BLOCO 1600
    ##### BLOCO 1600
    ##### BLOCO 1600
    #carrega registros 1600 do regerado
    R1600 = []
    bloco1600R = []
    q1600R = 0
    encR = comum.encodingDoArquivo(nomereg)
#    encR = encodingDoArquivo(nomereg)
    tempR = open(nomereg, 'r', encoding=encR, errors='ignore')
    for lR in tempR:
        if lR.startswith('|1600|') :
            q1600R = q1600R + 1
            bloco1600R.append(lR)
            cod = lR.split('|')[2]
            if (not cod in R1600):
                R1600.append(cod)
    tempR.close() 
    log("Quantidade de registros 1600 no REGERADO = ", q1600R)
    
    #verifica registros 1600 do protocolado
    P1600 = []
    bloco1600P = []
    q1600P = 0
    encP = comum.encodingDoArquivo(nomepro)
#    encP = encodingDoArquivo(nomepro)
    tempP = open(nomepro, 'r', encoding=encP, errors='ignore')
    for lP in tempP:
        if lP.startswith('|1600|') :
            q1600P = q1600P + 1
            bloco1600P.append(lP)
            cod = lP.split('|')[2]
            if (not cod in P1600):
                P1600.append(cod)
    tempP.close() 
    log("Quantidade de registros 1600 no PROTOCOLADO = ", q1600P)
    

    if enx1600 == 'N' and (aaaamm > '201707' and q1600P == 0 ):
        log('-'* 160)
        log('####')
        log("ERRO - ARQUIVO PROTOCOLADO NÃO POSSUI O REGISTRO 1600.", nomepro)
        log('####')
        log('-'* 160)
        ret = 1
    
    if( aaaamm < '201708' and q1600P > 0):
        log('-'* 160)
        log('####')
        log('ERRO - ARQUIVO PROTOCOLADO POSSUIA', q1600P , 'REGISTROS 1600 QUE FORAM SUBSTITUIDOS POR',q1600R,'QUE EXISTIAM NO REGERADO.')
        log('####')
        log('-'* 160)
        ret = 0
    
    if( (enx1600 == 'S'  or ( enx1600 == 'N'  and aaaamm < '201708' )) and q1600R == 0 ): 
        log('-'* 160)
        log('####')
        log("ERRO - O ARQUIVO REGERADO NÃO POSSUI REGISTROS 1600, O ENXERTO DO 1600 NÃO FOI REALIZADO.")
        log('####')
        log('-'* 160)
        ret = 1
    

    ##### BLOCO 0150
    ##### BLOCO 0150
    ##### BLOCO 0150
    ##### BLOCO 0150
    #carrega registros 1600
    bloco0150R = []
    q0150R = 0 

#    log("q1600R =", q1600R)
    if (q1600R > 0):

        registros31600 = []        
        for registros1600 in bloco1600R:
            reg_3_1600 = registros1600.split('|')[2]
            registros31600.append(reg_3_1600)

        tempR = open(nomereg, 'r', encoding=encR, errors='ignore')
        for lR in tempR:
            if lR.startswith('|0150|') :
                cod = lR.split('|')[2]

                if (cod in registros31600):
                    q0150R = q0150R + 1
                    bloco0150R.append(lR)
        tempR.close() 
    
#    print("Retorna ",bloco0150R,bloco1600R,ret)
#    input("continua 2?")

    




    return(bloco0150R,bloco1600R,ret)
    
    
def enxerto1400(enx,reg,pla):
        
    ret = 0
    #### - Carrega planilha 1400
    l1400   = []
    p1400   = load_workbook(os.path.join(pla))
    enxT    = os.path.join(enx+".temp")
    enx     = os.path.join(enx)
    aba1400 = p1400['REGISTRO 1400']
    
    for line in aba1400:
        v1400 = line[4].value

        if (v1400 != None):
            if(v1400[0] == '|'):
                l1400.append(v1400+'\n')
  
    q1400 = len(l1400)

    if (q1400 < 1):
        log("ERRO - Não existe registros 1400 a serem enxertados.")
        ret = 99
        return(ret)
        
    if (os.path.isfile(enx)):
        log(" Aguarde, enxertando bloco 1400: ",enx)
#        log("eold = ",eold)
#        log("enew = ",enew)
#        log("enx = ",enx)
#        log("pla = ",pla)
        
        
        shutil.move(enx,enxT, copy_function = shutil.copytree)
  

##### A PARTIR DAQUI, PROTOCOLADO PASSA A SER O ENXERTADO ANTIGO (enxT), E O ENXERTADO PASSA A SER O  (enx)
   
    path_nomeArquivoProtocolado = enxT
    path_nomeArquivoEnxertado   = enx
    
    
#Reabre arquivo PROTOCOLADO para processamento principal     
    encP = comum.encodingDoArquivo(path_nomeArquivoProtocolado)
#    encP = encodingDoArquivo(path_nomeArquivoProtocolado)
    arquivoP = open(path_nomeArquivoProtocolado, 'r', encoding=encP, errors='ignore')
    num_linhaP = 0

#Cria o arquivo enxertado 
    arquivoE = open(path_nomeArquivoEnxertado, 'w', encoding=encP)
    numlinE = 0


    #ler até encontrar o primeiro registro....
    for linhaP in arquivoP:
        num_linhaP = num_linhaP + 1
        if linhaP.startswith('|0000|') :
            break
    valorP = linhaP.split('|')[1]
    
    
   #grava do PROTOCOLADO até chegar no 1010
    while valorP != '1010':
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
        
    if (valorP == '1010'):
        nlP = ""
        pos = 0 
        for l in linhaP:
            if( pos == 14 ):
                if(q1400 > 0):
                    nlP = nlP + "S"
                else:    
                    nlP = nlP + "N"
            else:
                nlP = nlP + l
            pos = pos + 1
        numlinE = gravar(arquivoE, nlP[:], numlinE)
    else:        
        log("ERRO - Não existe registro 1010 no ENXERTADO 1400.")
        ret = 99
        return(ret)

    #enxerta as linhas da planilha registro_1400
    for linha1400 in l1400:
        numlinE = gravar(arquivoE, linha1400[:], numlinE)   
                  
    #posiciona o protocolado no > 1400
    while valorP <= '1400':
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
        

    #grava do PROTOCOLADO até achar o primeiro contador 9900
    while valorP != '9900':
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
    
    
#CONTAGEM DOS REGISTROS    
#CONTAGEM DOS REGISTROS    
#CONTAGEM DOS REGISTROS    
#CONTAGEM DOS REGISTROS    
#CONTAGEM DOS REGISTROS    

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
    arquivoE.close()

    if (os.path.isfile(enxT)):
        os.remove(enxT)

    return(ret)


def enxerto1600(pro, reg, enx ,b0150, b1600):
    ret = 0
    enxT  = os.path.join(enx+".temp")
    shutil.move(enx,enxT, copy_function = shutil.copytree)

##### A PARTIR DAQUI, PROTOCOLADO PASSA A SER O ENXERTADO TEMPORARIO (enxT), E O ENXERTADO NOVO PASSA A SER O ENXERTADO (enx)
   
    path_nomeArquivoProtocolado = enxT
    path_nomeArquivoEnxertado   = enx
    
    
#Reabre arquivo PROTOCOLADO para processamento principal     
    encP = comum.encodingDoArquivo(path_nomeArquivoProtocolado)
#    encP = encodingDoArquivo(path_nomeArquivoProtocolado)
    arquivoP = open(path_nomeArquivoProtocolado, 'r', encoding=encP, errors='ignore')
    num_linhaP = 0

#Cria o arquivo enxertado 
    arquivoE = open(path_nomeArquivoEnxertado, 'w', encoding=encP)
    numlinE = 0

    #ler até encontrar o primeiro registro....
    for linhaP in arquivoP:
        num_linhaP = num_linhaP + 1
        if linhaP.startswith('|0000|') :
            break
    valorP = linhaP.split('|')[1]
    
    #Gravar do PROTOCOLADO até chegar no 0150
    while valorP < '0150':
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break





    #grava os 0150 do regerado que existem no 1600 a ser inserido   
    if (len(b0150) > 0 and enx1600 == 'S'):
        
        for linha0150 in b0150:
            numlinE = gravar(arquivoE, linha0150[:], numlinE)
            
    #retorna ao PROTOCOLADO



            
   
    
   #grava do PROTOCOLADO até chegar no 1010
    while valorP != '1010':
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
        
    if (valorP == '1010'):
        nlP = ""
        pos = 0 
        for l in linhaP:
            if( pos == 18 and len(b1600) > 0 and enx1600 == 'S'):
                nlP = nlP + "S"
            else:
                nlP = nlP + l
            pos = pos + 1
        numlinE = gravar(arquivoE, nlP[:], numlinE)
    else:        
        log("ERRO - Não existe registro 1010 no ENXERTADO 1400.")
        ret = 99
        return(ret)


    #pula o 1010 que já existia
    while valorP == '1010':
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break


    #Gravar do PROTOCOLADO até chegar no 1600
    while valorP < '1600':
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break

    


    #enxerta o 1600
    #grava os 1600 do regerado a serem enxertados  
    if (len(b1600) > 0 and enx1600 == 'S'):
        
        for linha1600 in b1600:
            numlinE = gravar(arquivoE, linha1600[:], numlinE)
            
    #retorna ao PROTOCOLADO
    





                  
    #posiciona o protocolado no > 1600
    while valorP <= '1600':
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
        





    #grava do PROTOCOLADO até achar o primeiro contador 9900
    while valorP != '9900':
        numlinE = gravar(arquivoE, linhaP[:], numlinE)
        for linhaP in arquivoP:
            valorP = linhaP.split('|')[1]
            break
    
    
#CONTAGEM DOS REGISTROS    
#CONTAGEM DOS REGISTROS    
#CONTAGEM DOS REGISTROS    
#CONTAGEM DOS REGISTROS    
#CONTAGEM DOS REGISTROS    

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
    arquivoE.close()

    if (os.path.isfile(enxT)):
        os.remove(enxT)

    return(ret)




if __name__ == "__main__":
    global ret
    global series_troca
    
    log('#'*100)
    log("# ")  
    log("# - INICIO - ENXERTO_SPED")
    log("# ")
    log('#'*100)
    ret = 0
    retorno = parametros()
    nome1400 = "" 
    retproc =("0","","")
    

    if (retorno[0] != False):
        ret     = 0
        ufi     = retorno[0]
        mesanoi = retorno[1]
        mesi    = retorno[2]
        anoi    = retorno[3]
        iei     = retorno[4]
        enx1400 = retorno[5]
        enx1600 = retorno[6]
        series_troca = retorno[7]
        
        anomes = anoi+mesi
        
        #log(" = ", )
        #log("ret     = ", ret)
        #log("ufi     = ", ufi)
        #log("mesanoi = ", mesanoi)
        #log("mesi    = ", mesi)
        #log("anoi    = ", anoi)
        #log("iei     = ", iei)
        #log("enx1400 = ", enx1400)
        #log("enx1600 = ", enx1600)
        #log("anomes = ", anomes)
        #log(" = ", )
        
        if (enx1600=="N" and int(anomes) > 201707):
            e1600 = "N"
            enx1600 = "N"
        else:
            e1600 = "S"
            enx1600 = "S"
            
        if (enx1400 == 'S'):
            # Verifica se a planilha com os registros 1400 existe e pega o nome com o caminho completo.
            nome1400 = verifica1400(ufi,iei,mesi,anoi)
            if nome1400 == "":
                ret = 100






        if ( ret == 0):                    

            log("-"*100)
            log("# Processando ENXERTO SPED para os seguintes parâmetros:")
            log("#    UF   = ",ufi)
            log("#    MÊS  = ",mesi)
            log("#    ANO  = ",anoi)
            log("#    IE   = ",iei)
            log("#    1400 = ",enx1400)
            log("#    1600 = ",enx1600)
            log("#    ARQ. = ",nome1400)
            log("#    SERIE= ",series_troca)
        
            log("-"*100)
     
        
            retproc = processar(ufi,mesanoi,mesi,anoi,iei)

            ret  = retproc[0]
            nomeenx = retproc[1]
            nomereg = retproc[2]
            nomepro = retproc[3]
            
            
            
#        input("Vai realizar o 1400....")            
            
            
            
            
        if (enx1400 == "S" and nome1400 != "" and ret == 0):
            log("Realizando o enxerto do bloco 1400...")
            toti = {}
            totc = {}
            totais = {}
            toti['9900'] = 4
            ret = enxerto1400(nomeenx,nomereg,nome1400)    
            log("Fim do enxerto do bloco 1400...")
            
        if (ret == 0):
            log("Preparando para o enxerto dos blocos 0150 e 1600...")
            ret1600 = preparaenxerto1600(nomeenx,nomereg,nomepro,anoi+mesi,e1600)    
            b0150=ret1600[0]
            b1600=ret1600[1]  
            rp1600 = ret1600[2]
            log("Quantidade de registros 0150 = ",len(b0150))
            log("Quantidade de registros 1600 = ",len(b1600))
            log("Fim do preparo para o enxerto dos blocos 0150 e 1600...")
            ret = rp1600

    

#        input("Vai realizar o 1600....")            





        if ((e1600 == "S" and ret == 0) and (b1600 != [])):
            log("Realizando o enxerto do bloco 1600...")
            toti = {}
            totc = {}
            totais = {}
            toti['9900'] = 4
            ret = enxerto1600(nomepro, nomereg, nomeenx, b0150, b1600)
            log("Fim do enxerto do bloco 1600...")
        else:
            log("#### ATENCÃO: Enxerto dos blocos 0150 e 1600 Não foi realizado.")
            # ret = 1
            
        
    else:
        ret = 99
        
    log('#'*100)
    log("# ")  
    log("# - FIM - ENXERTO_SPED")
    log("# ")
    log("#"*100)

    log("Codigo de saida = ",ret)
    sys.exit(ret)




 



