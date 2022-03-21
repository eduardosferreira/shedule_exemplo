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
----------------------------------------------------------------------------------------------
"""
import os
import datetime
import atexit
import sys
import re
from pathlib import Path
from openpyxl import load_workbook
import shutil
global ret

msg = 0
toti = {}
totc = {}
totais = {}

toti['9900'] = 4

relatorio_erros = None

separadorDiretorio = ('/' if os.name == 'posix' else '\\')
SD=separadorDiretorio

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
            # print('Registrar :', args)
            self.__linhas_relatorio.append([x for x in args])
        else :
            print("#### ERRO ao registar dados no relatorio.")
            print(*args)
            raise
        return True
        

    def close(self):
        if len(self.__linhas_relatorio) > 0 :
            fd = open( self.__path_relatorio, 'w' )
            fd.write(';'.join( x.replace('\n', '') for x in self.__colunas_relatorio ) + '\n')
            for linha in self.__linhas_relatorio :
                fd.write( ';'.join( str(x).replace('\n', '').replace('\r', '') for x in linha ) + '\n' )
            fd.close()
            print('Gerado relatorio ..: %s'%(self.__arquivo_relatorio))


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
        for blc in self.blocos_paraCopia:
            if txt_linha.startswith('|' + blc + '|'):
                self.blocos_paraCopia[blc] += 1
                self.linhasBloco.append(txt_linha)
                encontrado = True
                break
        return encontrado

class Blocos_paraCopia:
    __blocos = None
    registroAncora = None

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
        for letraBloco in self.blocos:
            bloc = self.blocos[letraBloco]
            if bloc.naoTratarBloco: continue

            ## Checando se a linha pertence a um bloco que será sobrescrito
            if txt_linha.startswith('|'+letraBloco):
                for cod_bloco in bloc.blocos_paraCopia:
                    if txt_linha.startswith("|" + cod_bloco + "|"):
                        return True
                # bloc.qtd_linhasTotaisBloco += 1
                # print( "LIN", txt_linha)
                if txt_linha.startswith('|' + bloc.ancora_chave + '|'): ## Exemplo de linha : |D990|<qtde de linhas>
                    if not bloc.inserir_noFinal:
                        arq_novo.write(txt_linha)
                        bloc.inc_OutrasLinhasBlocoDestino()
                        print(formatNumero(num_linha), "Inserindo linhas no começo do bloco " + letraBloco)
                    else:
                        print(formatNumero(num_linha), "Inserindo linhas no final do bloco " + letraBloco)
                    
                    print('-'*70)
                    print('AKI - Qtde :', len(bloc.linhasBloco))
                    print('    - Qtde |9990 =', bloc.qtd_LinhasIniciadas9900)
                    print('    - Qtde |9    =', bloc.qtd_LinhasIniciadas9)
                    for line in bloc.linhasBloco :
                        arq_novo.write(line)
                        if line.startswith('|9990|'):
                            bloc.inc_Linhasiniciadas9900()
                        if line.startswith('|9'):
                            bloc.inc_Linhasiniciadas9()
                    print('    - Qtde |9990 =', bloc.qtd_LinhasIniciadas9900)
                    print('    - Qtde |9    =', bloc.qtd_LinhasIniciadas9)
                    print('-'*70)
                    # print('>>>>>>>>>\n', bloc.linhasBloco, '\n<<<<<<<<<<<<')

                    # arq_novo.writelines(bloc.linhasBloco)

                    if bloc.inserir_noFinal and txt_linha.startswith("|" +letraBloco+ "990|"):
                        print('Escreveu %s990 por AKI !!!! %s'%(letraBloco, bloc.qtd_OutrasLinhasBlocoDestino ))
                        arq_novo.write("|" + letraBloco + "990|" + str(bloc.qtd_linhasTotaisBloco+1) + "|\n")
                        bloc.inc_OutrasLinhasBlocoDestino()
                        return True

                    if not bloc.inserir_noFinal: return True
                elif txt_linha.startswith("|" +letraBloco+ "990|"):
                    print('FOI ESCRITO %s990 por AKI !!!!'%(letraBloco))
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
                        print(formatNumero(num_linha), "Inserindo linhas dos totalizadores do bloco "+letraBloco)
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

def encodingDoArquivo(path_arq) :
    global ret
    
    try :
        fd = open(path_arq, 'r', encoding='iso-8859-1')
        fd.read()
        fd.close()
    except :
        return 'utf-8'

    return 'iso-8859-1'

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
        print("-"*100)
        print('#### ERRO:    Arquivo %s não está na pasta %s'%(mascara,diretorio))
        print("-"*100)
    return(nomearq)

def validauf(uf):
    return(True if (uf.upper() in ('AC','AL','AM','AP','BA','CE','DF','ES','GO','MA','MG','MS','MT','PA','PB','PE','PI','PR','RJ','RN','RO','RR','RS','SC','SE','SP','TO')) else False)
          
def dtf():
    return (datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'))

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
            print("#   ",qdade, " => ", f, " IE = ", ie)
            try:
                ies.index(str(f).split("_")[4])
            except:
                ies.append(str(f).split("_")[4])
                continue
            
    else: 
        print('#### ERRO:    Arquivo %s não está na pasta %s'%(mascara,diretorio))
        ret=99
        return("")
    print("-"*100)
    return(ies)

def processar(ufi,mesanoi,mesi,anoi,iei):
    global ret
    
    nome_protocolado=""
    nome_regerado=""
    nome_enxertado=""
    dir_base = SD + 'arquivos' + SD + 'SPED_FISCAL'
    dir_protocolados = os.path.join(dir_base, 'PROTOCOLADOS', ufi, anoi, mesi)
    dir_regerados = os.path.join(dir_base, 'REGERADOS', ufi, anoi, mesi)
    dir_enxertados = os.path.join(dir_base, 'ENXERTADOS',ufi,anoi,mesi)
     
    mascara_regeradoi = "SPED_"+mesanoi+"_"+ufi+"_"+iei+"_REG*.txt"
    listadeies = ies_existentes(mascara_regeradoi,dir_regerados)
    
    #print(" = ", )
    #print(" = ", )
    #print(" = ", )
    #print("listadeies = ", listadeies)
    #print(" = ", )
    
    for iee in listadeies:
        
        print("#")
        print("#")
        print("#")
        print("-"*100)
        print("#")
        print("# ",dtf() , " INÍCIO do processamento para a IE ", iee)
        
        mascara_regerado = "SPED_"+mesanoi+"_"+ufi+"_"+iee+"_REG*.txt"
        mascara_protocolado = "SPED_"+mesanoi+"_"+ufi+"_"+iee+"_PROT*.txt"
        
        nome_regerado = nome_arquivo(mascara_regerado,dir_regerados)
        nome_protocolado = nome_arquivo(mascara_protocolado,dir_protocolados)
         
        if ((nome_regerado == "") or (nome_protocolado == "")):
            print("-"*100)
            print("#### ERRO:    Não foi processado o ENXERTO para a dupla de arquivos:")
            print("#### ERRO:        Arquivo regerado    = ", nome_regerado)
            print("#### ERRO:        Arquivo protocolado = ", nome_protocolado)
            print("-"*100)
            ret=99
        else:
                
            ### prepara saida ENXERTADO

            if (str(nome_regerado).count("_") == 6):
                versao_enxertado = "_"+(str(nome_regerado).split(".")[0]).split("_")[6]
            else:
                versao_enxertado = ""
            nome_enxertado = os.path.join(dir_enxertados, "SPED_"+mesanoi+"_"+ufi+"_"+iee+"_ENX"+versao_enxertado+ ".txt")
        
            if not os.path.isdir(dir_enxertados) :
                os.makedirs(dir_enxertados)

            print("#")
            print("-"*100)
            print("#  Arquivos a serem processados:")
            print("#     Arquivo protocolado = ", nome_protocolado)
            print("#     Arquivo regerado    = ", nome_regerado)
            print("#     Arquivo enxertado   = ", nome_enxertado)
            print("-"*100)
            
            if (processaDiretorio(nome_protocolado, nome_regerado, nome_enxertado, dir_protocolados, dir_regerados, dir_enxertados) == False):
                ret = 99
                
        print("-"*100)
        print("#")
        print("# ",dtf() , " FIM do processamento para a IE ", iee)
        print("#")
        print("-"*100)
        print("#")
        print("#")
        print("#")

    return(ret,nome_enxertado,nome_regerado,nome_protocolado)

def parametros():
    global ret
    ufi = "SP"
    mesanoi = ""
    iei = "*"
    mesi = ""
    anoi = "" 
    ret = 0
    
#### Recebe, verifica e formata os argumentos de entrada.  
    enx1400 = 'S'
    if (len(sys.argv) == 5
        and validauf(ufi)
        and len(sys.argv[1])==6  
        and int(sys.argv[1][:2])>0 
        and int(sys.argv[1][:2])<13
        and int(sys.argv[1][2:])<=datetime.datetime.now().year
        and int(sys.argv[1][2:])>2014
        and (sys.argv[2].upper() in ("S", "'S'", '"S"',"N", "'N'", '"N"'))
        ):
     
        mesanoi = sys.argv[1].upper()

        if len(sys.argv) == 5:
            iei=sys.argv[4].upper()
            iei = re.sub('[^0-9]','',iei)
       
            if ( (iei == "") or (iei == "''") or (iei == '""') or (int("0"+iei) == 0)):
                iei = "*"
            
    else :
        print("-" * 100)
        print("#### ")
        print('#### ERRO - Erro nos parametros do script.')
        print("#### ")
        print('#### Exemplo de como deve ser :')
        print('####      %s <MMAAAA> <S/N> <s/n> <IE>'%(sys.argv[0] if sys.argv[0][0] == '.' else '.' + SD + sys.argv[0] ))
        print("#### ")
        print('#### Onde')
        # print('####      <UF> = estado. Ex: SP')
        print('####      <MMAAAA> = mês e ano. Ex: Para junho de 2020 informe 062020.')
        print('####      <S/N>    = Deve ser informado S para enxertar o bloco 1400 ou N para que não tenha o bloco 1400.')
        print('####      <s/n>    = Deve ser informado S para enxertar o bloco 1600 regerado independente da data de origem ')
        print('####                     ou N para que considere as regras de data até 072017 e registros 0150.')
        print('####      <IE>     = Inscição Estadual. É opcional, pode ou não ser informado.')
        print('####                     caso não informado, será processado para todas IEs do estado <UF> informado.')
        print("#### ")
        # print('#### Portanto, se o estado = SP, o mes = 06 e o ano = 2020, e deseja todas IEs,  o comando correto deve ser :')  
        print('#### Portanto, se o mes = 06 e o ano = 2020, deseja enxertar o bloco 1400 , bloco 1600 e IE = 108383949112,  o comando correto deve ser :')  
        print('####      %s 062020 S S 108383949112'%(sys.argv[0] if sys.argv[0][0] == '.' else '.' + SD + sys.argv[0]))  
        print("#### ")
        print("-" * 100)
        print("")
        print("Retorno = 99") 
        ret = 99

        return(False,False,False,False,False,False)

    mesanoi = sys.argv[1].upper()
    mesi    = sys.argv[1][:2].upper()
    anoi    = sys.argv[1][2:].upper()
    enx1400 = sys.argv[2].upper()
    enx1600 = sys.argv[3].upper()
    
    return(ufi,mesanoi,mesi,anoi,iei,enx1400,enx1600)

def processaDiretorio(nome_protocolado, nome_regerado, nome_enxertado, path_protocolados, path_regerados, path_enxertados) :
    global relatorio_erros 
    
    path_nomeArquivoOriginal = nome_protocolado
    path_nomeArquivoBlocos = nome_regerado
    path_nomeArquivoNovoArq = nome_enxertado

    blocos_paraCopia = Blocos_paraCopia(
        'D695', 'D696', 'D697'
    )

    blocos_paraCopia.blocos['D'].inserir_noFinal = True

    # blocos_paraCopia.blocos['E'].inserir_noFinal = False
    # blocos_paraCopia.blocos['E'].ancora_chave = 'E100'
    # blocos_paraCopia.blocos['E'].naoTratarBloco = True

    encoding = encodingDoArquivo(path_nomeArquivoBlocos)
    arquivoCopia = open(path_nomeArquivoBlocos, 'r', encoding=encoding)
    reg_inicial_arqCopia = False
    for linha in arquivoCopia:
        if linha.startswith('|0000|') :
            reg_inicial_arqCopia = linha[:]
        blocos_paraCopia.processarLinha_Origem(linha)

    arquivoCopia.close()

    print("Registros selecionados no arquivo ", path_nomeArquivoBlocos)

    for bloco in blocos_paraCopia.blocos:
        print("Bloco:", bloco)
        bloco = blocos_paraCopia.blocos[bloco]
        for chave in bloco.blocos_paraCopia:
            print("  ", chave+':', formatNumero(bloco.blocos_paraCopia[chave]))

    print("Contando linhas do arquivo original...")
    qtdLinhasArqOriginal = contarLinhasArquivo(path_nomeArquivoOriginal)

    print("Abrindo arquivo original e escrevendo arquivo novo...")
    encoding = encodingDoArquivo(path_nomeArquivoOriginal)
    arquivoOriginal = open(path_nomeArquivoOriginal, 'r', encoding=encoding, errors='ignore')
    # arquivoNovo = open(path_nomeArquivoNovoArq, 'w')
    arquivoNovo = open(path_nomeArquivoNovoArq, 'w', encoding=encoding, errors='ignore')

    num_linhaAtual = 0
    num_linhaOriginaisCopiadas = 0
    qtde_registros_9900 = 0 #### Conta todos os registros iniciados com |9900|
    qtde_registros_9990 = 0 #### Conta todos os registros iniciados com |9 + o ultimo registro |9999|

    for linha in arquivoOriginal:
        num_linhaAtual += 1
        if num_linhaAtual == 1 and reg_inicial_arqCopia :
            if (linha.split('|')[4] != reg_inicial_arqCopia.split('|')[4]) or (linha.split('|')[5] != reg_inicial_arqCopia.split('|')[5]) :
                print("#"*80)
                print('### ERRO - Arquivos com periodos de dados diferentes ... Verifique !!')
                print("#"*80)
                relatorio_erros.registrar( 'ERRO', 'Arquivos com periodos de dados diferentes.' )
                return False

        if num_linhaAtual % 500000 == 0:
            print(formatNumero(num_linhaAtual), "/", formatNumero(qtdLinhasArqOriginal))

        if not blocos_paraCopia.processarLinha_Destino(linha, num_linhaAtual, arquivoNovo):
            if linha.startswith("|9999|"):
                qtd_linhas = num_linhaOriginaisCopiadas + blocos_paraCopia.qtd_linhas_geradas_blocos + 1 #|9999|
                print("Escrevendo registro 9999: "+formatNumero(qtd_linhas))
                arquivoNovo.write("|9999|" + str(qtd_linhas) + "|\n")
                break
            else:
                linha_de_bloco = False
                for bloco in blocos_paraCopia.blocos :
                    if linha.startswith('|' + bloco) :
                        linha_de_bloco = bloco
                
                ##### Realizar nova somatoria dos registros |9900|

                if linha.startswith("|9900|9900"):
                    print(">>> ANTES DA SOMA >> Quantidade de registros |9900| =", qtde_registros_9900)
                    for bloco in blocos_paraCopia.blocos :
                        qtde_registros_9900 += blocos_paraCopia.blocos[bloco].qtd_LinhasIniciadas9900
                    qtde_registros_9900 += 1
                    qtde_registros_9990 += 1 ##### Como esse registro inicia com |9 tem que incrementar o |9990|
                    print(">> Quantidade de registros |9900| =", qtde_registros_9900 )
                    linha = "|9900|9900|" + str(qtde_registros_9900) + "|\n"
                                
                ##### Realizar nova somatoria dos registros |9*
                elif linha.startswith("|9990|"):
                    print(">>> ANTES DA SOMA >> Quantidade de registros |9 =", qtde_registros_9990 )
                    for bloco in blocos_paraCopia.blocos :
                        qtde_registros_9990 += blocos_paraCopia.blocos[bloco].qtd_LinhasIniciadas9
                    qtde_registros_9990 += 1
                    print(">> Quantidade de registros |9 =", qtde_registros_9990 )
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

    print("Fechando arquivos")

    for bloco in blocos_paraCopia.blocos :
        print("Linhas |9 =", blocos_paraCopia.blocos[bloco].qtd_LinhasIniciadas9)
        print("Linhas |9900 =", blocos_paraCopia.blocos[bloco].qtd_LinhasIniciadas9900)

    arquivoOriginal.close()
    arquivoNovo.close()


    #####    encoding = encodingDoArquivo(path_nomeArquivoOriginal)  - Roney


    arquivoNovo = open(path_nomeArquivoNovoArq, 'r', encoding=encoding, errors='ignore')
    ultimas_linhas = ['','']
    for lin in arquivoNovo.readlines() :
        ultimas_linhas[1] = ultimas_linhas[0]
        ultimas_linhas[0] = lin
    arquivoNovo.close()

    if not ultimas_linhas[0].startswith('|9999|') :
        if not ultimas_linhas[1].startswith('|9999|') :
            print("#"*80)
            print("### Erro arquivo 'Enxertado' não possui a ultima linha com registro |9999|")
            print("#"*80)
            relatorio_erros.registrar( 'ERRO', 'Arquivo ENXERTADO não possui a ultima linha com registro |9999|' )
            return False
    
    return True





def verifica1400(uf1400,ie1400,mes1400,ano1400):
    pasta_1400 = SD + 'arquivos' + SD + 'REGISTRO_1400' + SD + 'RELATORIOS' + SD + uf1400 + SD + str(ano1400) + SD + str(mes1400) + SD
    mask_1400  = ie1400+"_"+"Valores_Agregados_1400_"+ mesanoi+".xlsx"
    nome_1400  = nome_arquivo(mask_1400,pasta_1400)
    if (nome_1400 == ""):
        print("#### ERRO - Execute antes o script registro_1400 para este mes, ano e IE. ")
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
#        print("Nova chave toti = ", chaveti)
        toti[chaveti] = 1
        
    #total por chave
    if (chave in totc):
        totc[chave] = totc[chave] + 1
    else:
#        print("Nova chave totc = ", chave)
        totc[chave] = 1

    if ((chave[1:] == '990' )):
        dado = '|'+ str(chaveti) + '|'+ str(toti[chaveti]) + '|\n'

    #grava no destino
    arquivo.write(dado)
    if contador % 500000 == 0:
        print(formatNumero(contador))
    
    return (contador)

def preparaenxerto1600(nomeenx, nomereg, nomepro, aaaamm, enx1600):
    global msg
    msg = 0
    
    #print("Recebeu nomeenx = ", nomeenx )
    #print("Recebeu nomereg = ", nomereg )
    #print("Recebeu nomepro = ", nomepro )
    #print("Recebeu aaaamm  = ", aaaamm )
    #print("Recebeu enx1600 = ", enx1600 )

   
    ##### BLOCO 1600
    ##### BLOCO 1600
    ##### BLOCO 1600
    ##### BLOCO 1600
    #carrega registros 1600 do regerado
    R1600 = []
    bloco1600R = []
    q1600R = 0
    encR = encodingDoArquivo(nomereg)
    tempR = open(nomereg, 'r', encoding=encR, errors='ignore')
    for lR in tempR:
        if lR.startswith('|1600|') :
            q1600R = q1600R + 1
            bloco1600R.append(lR)
            cod = lR.split('|')[2]
            #print("cod = ", cod)
            if (not cod in R1600):
                R1600.append(cod)
    tempR.close() 
    
    #verifica registros 1600 do protocolado
    P1600 = []
    bloco1600P = []
    q1600P = 0
    encP = encodingDoArquivo(nomepro)
    tempP = open(nomepro, 'r', encoding=encP, errors='ignore')
    for lP in tempP:
        if lP.startswith('|1600|') :
            q1600P = q1600P + 1
            bloco1600P.append(lP)
            cod = lP.split('|')[2]
            if (not cod in P1600):
                P1600.append(cod)
    tempP.close() 

    if(aaaamm > '201707' and q1600P == 0 ):
        print('-'* 160)
        print('####')
        print('#### - ERRO - ARQUIVO PROTOCOLADO NÃO POSSUI O REGISTRO 1600.', nomepro)
        print('####')
        print('-'* 160)
        msg = 1
    
    if( aaaamm < '201708' and q1600P > 0):
        print('-'* 160)
        print('####')
        print('#### - ERRO - ARQUIVO PROTOCOLADO POSSUIA', q1600P , 'REGISTROS 1600 QUE FORAM SUBSTITUIDOS POR',q1600R,'QUE EXISTIAM NO REGERADO.')
        print('####')
        print('-'* 160)
        msg = 1
    
    if( (enx1600 == 'S'  or ( enx1600 == 'N'  and aaaamm < '201708' )) and q1600R == 0 ): 
        print('-'* 160)
        print('####')
        print('#### - ERRO - O ARQUIVO REGERADO NÃO POSSUI REGISTROS 1600, O ENXERTO DO 1600 ÃO FOI REALIZADO.')
        print('####')
        print('-'* 160)
        msg = 1
    

    ##### BLOCO 0150
    ##### BLOCO 0150
    ##### BLOCO 0150
    ##### BLOCO 0150
    #carrega registros 1600
    bloco0150R = []
    q0150R = 0 





    #print("q1600R     =", q1600R)
    #print("nomereg    =", nomereg)
    #print("bloco1600R = ", bloco1600R)
    #print("R1600      = ", R1600)





    if (q1600R > 0):

        tempR = open(nomereg, 'r', encoding=encR, errors='ignore')
        for lR in tempR:
            if lR.startswith('|0150|') :
                cod = lR.split('|')[2]
                #print("cod = ", cod)
                if (cod in R1600):
                    #print("lR = ", lR)
                    q0150R = q0150R + 1
                    bloco0150R.append(lR)
        tempR.close() 
    
    #print("RETORNO = bloco0150R = ", bloco0150R)
    #print("RETORNO = bloco1600R = ", bloco1600R)
    
    return(bloco0150R,bloco1600R)
    
    
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
        print("#### ERRO - Não existe registros 1400 a serem enxertados.")
        ret = 99
        return(ret)
        
    if (os.path.isfile(enx)):
        print(" Aguarde, enxertando bloco 1400: ",enx)
#        print("eold = ",eold)
#        print("enew = ",enew)
#        print("enx = ",enx)
#        print("pla = ",pla)
        
        
        shutil.move(enx,enxT, copy_function = shutil.copytree)
  

##### A PARTIR DAQUI, PROTOCOLADO PASSA A SER O ENXERTADO ANTIGO (enxT), E O ENXERTADO PASSA A SER O  (enx)
   
    path_nomeArquivoProtocolado = enxT
    path_nomeArquivoEnxertado   = enx
    
    
#Reabre arquivo PROTOCOLADO para processamento principal     
    encP = encodingDoArquivo(path_nomeArquivoProtocolado)
    arquivoP = open(path_nomeArquivoProtocolado, 'r', encoding=encP, errors='ignore')
    num_linhaP = 0

#Cria o arquivo enxertado 
    # arquivoE = open(path_nomeArquivoEnxertado, 'w')
    arquivoE = open(path_nomeArquivoEnxertado, 'w', encoding=encP, errors='ignore')
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
        print("#### ERRO - Não existe registro 1010 no ENXERTADO 1400.")
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
#    print("linhaf = ", linhaf)

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
    encP = encodingDoArquivo(path_nomeArquivoProtocolado)
    arquivoP = open(path_nomeArquivoProtocolado, 'r', encoding=encP, errors='ignore')
    num_linhaP = 0

#Cria o arquivo enxertado 
    # arquivoE = open(path_nomeArquivoEnxertado, 'w')
    arquivoE = open(path_nomeArquivoEnxertado, 'w', encoding=encP, errors='ignore')
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
#    if (len(b0150) > 0 and enx1600 == 'S'):
    if (len(b0150) > 0):
        
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
#            if( pos == 18 and len(b1600) > 0 and enx1600 == 'S'):
            if( pos == 18 and len(b1600) > 0 ):
                nlP = nlP + "S"
            else:
                nlP = nlP + l
            pos = pos + 1
        numlinE = gravar(arquivoE, nlP[:], numlinE)
    else:        
        print("#### ERRO - Não existe registro 1010 no ENXERTADO 1400.")
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
#    if (len(b1600) > 0 and enx1600 == 'S'):
    if (len(b1600) > 0):
        
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
#    print("linhaf = ", linhaf)

    arquivoP.close()
    arquivoE.close()

    if (os.path.isfile(enxT)):
        os.remove(enxT)

    return(ret)




if __name__ == "__main__":
    global ret
    
    print('#'*100)
    print("# ")  
    print("# ",dtf() , " - INICIO - ENXERTO_SPED")
    print("# ")
    print('#'*100)
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
        
        anomes = anoi+mesi
        
        #print(" = ", )
        #print("ret     = ", ret)
        #print("ufi     = ", ufi)
        #print("mesanoi = ", mesanoi)
        #print("mesi    = ", mesi)
        #print("anoi    = ", anoi)
        #print("iei     = ", iei)
        #print("enx1400 = ", enx1400)
        print("enx1600 = ", enx1600)
        print("anomes = ", anomes)
        #print(" = ", )
        
        
        
        if (enx1600=="N" and int(anomes) > 201707):
            e1600 = "N"
        else:
            e1600 = "S"
            enx1600 = "S"            
            
   
    
        if (enx1400 == 'S'):
            # Verifica se a planilha com os registros 1400 existe e pega o nome com o caminho completo.
            nome1400 = verifica1400(ufi,iei,mesi,anoi)
            if nome1400 == "":
                ret = 100
                


        if ( ret == 0):                    

            print("-"*100)
            print("# Processando ENXERTO SPED para os seguintes parâmetros:")
            print("#    UF   = ",ufi)
            print("#    MÊS  = ",mesi)
            print("#    ANO  = ",anoi)
            print("#    IE   = ",iei)
            print("#    1400 = ",enx1400)
            print("#    1600 = ",enx1600)
            print("#    ARQ. = ",nome1400)
        
            print("-"*100)
     
        
            retproc = processar(ufi,mesanoi,mesi,anoi,iei)

            ret  = retproc[0]
            nomeenx = retproc[1]
            nomereg = retproc[2]
            nomepro = retproc[3]
  
        if (enx1400 == "S" and nome1400 != "" and ret == 0):
            print("Realizando o enxerto do bloco 1400...")
            toti = {}
            totc = {}
            totais = {}
            toti['9900'] = 4
            ret = enxerto1400(nomeenx,nomereg,nome1400)    
            print("Fim do enxerto do bloco 1400...")
            
            
        if (ret == 0):
            print("Preparando para o enxerto dos blocos 0150 e 1600...")
            ret1600 = preparaenxerto1600(nomeenx,nomereg,nomepro,anoi+mesi,e1600)    
            b0150=ret1600[0]
            b1600=ret1600[1]    
            print("Quantidade de registros 0150 = ",len(b0150))
            print("Quantidade de registros 1600 = ",len(b1600))
            print("Fim do preparo para o enxerto dos blocos 0150 e 1600...")
            
  
        if ((e1600 == "S" and ret == 0) and (b1600 != [])):
            print("Realizando o enxerto do bloco 1600...")
            toti = {}
            totc = {}
            totais = {}
            toti['9900'] = 4
            
            
            ret = enxerto1600(nomepro, nomereg, nomeenx, b0150, b1600)
            print("Fim do enxerto do bloco 1600...")
        else:
            print("#### ATENCÃO: Enxerto dos blocos 0150 e 1600 Não foi realizado.")
            
        
    else:
        ret = 99
        
    print('#'*100)
    print("# ")  
    print("# ",dtf() , " - FIM - ENXERTO_SPED")
    print("# ")
    print("#"*100)

    ret = ret + msg

    print("Codigo de saida = ",ret)
    sys.exit(ret)







