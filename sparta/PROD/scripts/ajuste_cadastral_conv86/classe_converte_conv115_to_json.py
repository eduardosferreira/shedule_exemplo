#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
MODULO ...: TESHUVA
SCRIPT ...: retorna_dados_banco.py
CRIACAO ..: 11/02/2022
AUTOR ....: Victor Santos Cardoso - KYROS TECNOLOGIA
DESCRICAO.: 
----------------------------------------------------------------------------------------------
  HISTORICO : 
----------------------------------------------------------------------------------------------
"""
# from asyncio import constants
# from base64 import decode
# from distutils.log import Log
# from operator import truediv
import os
import sys

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)

import configuracoes
import comum
import layout

log.gerar_log_em_arquivo = True
comum.carregaConfiguracoes(configuracoes)
layout.carregaLayout()

class CONVERSOR :
    arquivo_item_aberto      = ''
    registro_item_corrente   = ''
    nome_arquivo_item_aberto = ''

    def tipoArquivo(self, path_arq) :
        try :
            fd = open(path_arq, 'r', encoding='iso-8859-1')
            fd.readline()
            fd.close()
        except :
            return 'utf-8'
        return 'iso-8859-1'

    def leArquivoMestre(self, diretorio_arquivos, ano_referencia, arquivo_controle, lista_notas, ini_volume, fim_volume):

        log('INICIANDO A LEITURA DO ARQUIVO')
        if int(ano_referencia) >= 2017:

            pos_tipo = 28
            versao   = ''

        else:
            pos_tipo = 10
            versao   = '_Antigo'
        
        arquivo_mestre = str(arquivo_controle)
        arquivo_mestre = arquivo_mestre[:pos_tipo] + 'M' + arquivo_mestre[pos_tipo + 1:]

        arquivo_destinatário = str(arquivo_controle)
        arquivo_destinatário = arquivo_destinatário[:pos_tipo] + 'D' + arquivo_destinatário[pos_tipo + 1:]

        arquivo_item = str(arquivo_controle)
        arquivo_item = arquivo_item[:pos_tipo] + 'I' + arquivo_item[pos_tipo + 1:]

        if not os.path.isfile(os.path.join(diretorio_arquivos,arquivo_mestre)):
            log(' ERRO,  Não foi encontrado o arquivo ',arquivo_mestre.center(100,'='))
            raise 'ERRO'

        if not os.path.isfile(os.path.join(diretorio_arquivos,arquivo_destinatário)):
            log(' ERRO,  Não foi encontrado o arquivo ',arquivo_destinatário.center(100,'='))
            raise 'ERRO'

        if not os.path.isfile(os.path.join(diretorio_arquivos,arquivo_item)):
            log(' ERRO,  Não foi encontrado o arquivo ',arquivo_item.center(100,'=') ) 
            raise 'ERRO'

        arq_mestre = open(os.path.join(diretorio_arquivos,arquivo_mestre), 'rb')
        primeira_linha = arq_mestre.readline()
        somador = len(primeira_linha) - 425
        arq_mestre.close()
        arq_mestre = open(os.path.join(diretorio_arquivos,arquivo_mestre), 'r', encoding=self.tipoArquivo(os.path.join(diretorio_arquivos,arquivo_mestre)) )
        
        retorno_mestre = []
        
        log('ARQUIVO MESTRE       ',arquivo_mestre) 
        log('ARQUIVO DESTINATARIO ',arquivo_destinatário) 
        log('ARQUIVO ITEM         ',arquivo_item)
        log(''.center(100,'='))
        
        contador_notas = 0

        for nota,num_item in lista_notas:
            
            if not (int(nota) >= ini_volume and int(nota) <= fim_volume):
                continue

            linha = int(nota) - ini_volume
            pulo = linha * (425 + somador )

            arq_mestre.seek(pulo, 0)

            linha_mestre = arq_mestre.readline()

            registro_mestre_quebrado = layout.quebraRegistroDicionario(linha_mestre, 'LayoutMestre' + versao)

            chave_nota = registro_mestre_quebrado['NUMERO_NF']
            
            if chave_nota == nota or not lista_notas:             
                
                contador_notas += 1
                
                if (contador_notas % 5000) == 0:

                    log('[ %s ] NOTAS PROCESSADAS ATÉ O MOMENTO !!!'%contador_notas)

                linha_item = int(registro_mestre_quebrado['REF_ITEM_NF']) + int(num_item) - 2

                registro_mestre_quebrado['item'] = self.leArquivoItem( diretorio_arquivos, ano_referencia, arquivo_item, registro_mestre_quebrado, linha_item, num_item )
                
                retorno_mestre.append(registro_mestre_quebrado)

        log('[ %s ] NOTAS PROCESSADAS POR THREAD!!!'%contador_notas)        
        arq_mestre.close()

        return retorno_mestre


    def leArquivoItem( self, diretorio_arquivos, ano_referencia, nome_arquivo_item, registro_mestre, linha_item, num_item ):

        if int(ano_referencia) >= 2017:
            versao = ''
        else:
            versao = '_Antigo'

        if nome_arquivo_item != self.nome_arquivo_item_aberto:

            self.arquivo_item_aberto = open(os.path.join(diretorio_arquivos, nome_arquivo_item), 'rb')
            
            primeira_linha = self.arquivo_item_aberto.readline()
            
            self.somador_item = len(primeira_linha) - 331
            
            self.arquivo_item_aberto.close()            

            self.arquivo_item_aberto = open(os.path.join(diretorio_arquivos, nome_arquivo_item), 'r', encoding= self.tipoArquivo(os.path.join(diretorio_arquivos,nome_arquivo_item)) )

            self.registro_item_corrente = self.arquivo_item_aberto.readline()

            self.nome_arquivo_item_aberto = nome_arquivo_item

        pulo = linha_item * (331 + self.somador_item )
        
        retorna_itens = []

        encontrou = False

        try:    
            self.arquivo_item_aberto.seek(pulo, 0)
            
            self.registro_item_corrente = self.arquivo_item_aberto.readline()

        except Exception as e:

            print('ERRO ', e)

            self.arquivo_item_aberto = open(os.path.join(diretorio_arquivos, nome_arquivo_item), 'r', encoding= self.tipoArquivo(os.path.join(diretorio_arquivos,nome_arquivo_item)) )
            
            self.arquivo_item_aberto.seek(pulo, 0)

            self.registro_item_corrente = self.arquivo_item_aberto.readline()

        if not encontrou and self.registro_item_corrente:  
        
            registro_item_corrente_quebrado = layout.quebraRegistroDicionario(self.registro_item_corrente, 'LayoutItem' + versao)

            encontrou_nota = True

            for chave in ['NUMERO_NF', 'DATA_EMISSAO', 'SERIE', 'MODELO']:
                
                chave_mestre = chave
                chave_item   = chave

                if registro_mestre[chave_mestre] != registro_item_corrente_quebrado[chave_item]:
                    
                    encontrou_nota = False
            
            if encontrou_nota:

                if registro_mestre['NUMERO_NF'] == registro_item_corrente_quebrado['NUMERO_NF'] and self.registro_item_corrente\
                      and num_item == registro_item_corrente_quebrado['NUM_ITEM']:
                    
                    registro_item_corrente_quebrado = layout.quebraRegistroDicionario(self.registro_item_corrente, 'LayoutItem' + versao)

                    retorna_itens.append(registro_item_corrente_quebrado)

                    self.registro_item_corrente = self.arquivo_item_aberto.readline()

                    registro_item_corrente_quebrado = layout.quebraRegistroDicionario(self.registro_item_corrente, 'LayoutItem' + versao)

                encontrou = True

            else:

                self.registro_item_corrente = self.arquivo_item_aberto.readline()

        if not self.registro_item_corrente:

            if self.arquivo_item_aberto:
            
                self.arquivo_item_aberto.close() 
                self.arquivo_item_aberto = False
        
        return retorna_itens

    def converte_conv115_to_json( self, diretorio_arquivos, ano_referencia, volume = 0, lista_notas = [], ini_volume = 0, fim_volume = 0):

        log('INICIANDO PROCESSAMENTO DO CONVERT_CONV115_TO_JSON')
        lista_arquivos = os.listdir(diretorio_arquivos)

        if int(ano_referencia) >= 2017:

            pos_tipo             = 28
            tamanho_nome_arquivo = 33
            versao               = ''

        else:

            pos_tipo             = 10
            tamanho_nome_arquivo = 15
            versao               = '_Antigo'

        lista_aprocessar = []

        for arquivo in lista_arquivos:

            if len(arquivo) == tamanho_nome_arquivo:

                if arquivo[pos_tipo] == 'C':

                    if int(volume) > 0:

                        split_volume = arquivo.split('.')

                        if split_volume[1] != volume:
                            continue
                else:
                    continue
            else:
                continue
            
            lista_aprocessar.append(arquivo)

        dados_por_volume = []  

        for arquivo in lista_aprocessar:

            arquivo_controle = open(os.path.join(diretorio_arquivos,arquivo), 'r', encoding='iso-8859-1')

            log('INICIANDO A LEITURA DO ARQUIVO DE CONTROLE: ')
            log('CONTROLE: ',arquivo)
            log(''.center(100,'='))

            for registro in arquivo_controle:
                
                registro_quebrado = layout.quebraRegistroDicionario(registro, 'LayoutControleV3' + versao)
                registro_quebrado['mestre'] = self.leArquivoMestre( diretorio_arquivos, ano_referencia, arquivo, lista_notas, ini_volume, fim_volume )
                        
                dados_por_volume.append(registro_quebrado)
                
        return dados_por_volume

if __name__ == "__main__" :
    
    ret = 0
    
    comum.addParametro( 'DIRETORIO_ARQUIVOS', None, "diretorio_arquivos" , True , '' )
    comum.addParametro( 'ANO_REFERENCIA'    , None, "ano_referencia"     , True , '')
    comum.addParametro( 'VOLUME'            , None, "volume"             , False , '')

    if not comum.validarParametros() :
        log('### ERRO AO VALIDAR OS PARÂMETROS')
        ret = 91
    else:
        configuracoes.diretorio_arquivos  = comum.getParametro('DIRETORIO_ARQUIVOS')
        configuracoes.ano_referencia      = comum.getParametro('ANO_REFERENCIA')
        configuracoes.volume              = comum.getParametro('VOLUME')

        conversor = CONVERSOR()
        if not conversor.converte_conv115_to_json(configuracoes.diretorio_arquivos, configuracoes.ano_referencia, configuracoes.volume) :
            log('ERRO no processamento do relatorio !')
            ret = 92

    sys.exit(ret)
