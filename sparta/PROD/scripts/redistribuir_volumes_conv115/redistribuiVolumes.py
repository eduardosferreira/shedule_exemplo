#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: GF
  MODULO ...: 
  SCRIPT ...: redistribuiVolumes.py
  CRIACAO ..: 07/01/2020
  AUTOR ....: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO : Processo para redistribuição dos volumes de arquivos Conv115 regerados, de forma 
              que os mesmos fiquem alinhados com a distribuição realizada no Protocolado.
              
----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 07/01/2020 - Welber Pena de Sousa - Kyros Tecnologia
            - Criacao do script.
    * 29/02/2020 - Flavio Teixeira - ALT001
            - Valida se diretorio existe antes de realizar a exclusao.
    * 06/04/2020 - Flavio Teixeira 
            - Incluindo diretoro protocolado NAS com Uf (variaveis['uf'])
    * 02/03/2022 - Eduardo da Silva Ferreira - Kyros Tecnologia
                 - [PTITES-1639] Padrão de diretórios do SPARTA            
----------------------------------------------------------------------------------------------
"""

import os
import sys

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)
import configuracoes

import shutil
import datetime
import hashlib

import comum
import sql
import layout
import util

name_script = os.path.basename(__file__).split('.')[0]
log.gerar_log_em_arquivo = True

comum.carregaConfiguracoes(configuracoes)

dic_registros = {}
dic_layouts = layout.carregaLayout()
dic_campos = {}
variaveis = {}
dic_fd = {}

#- [PTITES-1639]
configuracoes.gv_ds_diretorio_saida   = os.path.dirname(configuracoes.dir_geracao_arquivos)
configuracoes.gv_ds_diretorio_entrada = os.path.dirname(configuracoes.dir_entrada)

configuracoes.gv_ds_diretorio_ready = os.path.join(configuracoes.gv_ds_diretorio_saida, 'ready')
#if not os.path.isdir(configuracoes.gv_ds_diretorio_ready) and os.path.isdir(os.path.join( getattr(configuracoes,'pasta_trabalho', '.'), 'ready')):
#    configuracoes.gv_ds_diretorio_ready = os.path.join( getattr(configuracoes,'pasta_trabalho', '.'), 'ready')
configuracoes.gv_ds_diretorio_work = os.path.join(configuracoes.gv_ds_diretorio_saida, 'work')
#if not os.path.isdir(configuracoes.gv_ds_diretorio_work) and os.path.isdir(os.path.join( getattr(configuracoes,'pasta_trabalho', '.'), 'work')):
#    configuracoes.gv_ds_diretorio_work = os.path.join( getattr(configuracoes,'pasta_trabalho', '.'), 'work')

configuracoes.gv_ds_pasta_base_serie = os.path.join(configuracoes.gv_ds_diretorio_entrada, 'Protocolado')
#if not os.path.isdir(configuracoes.gv_ds_pasta_base_serie) and os.path.isdir(getattr(configuracoes,'pasta_base_serie', 'NAO_EXISTE_DIRETORIO_PROTOCOLADO')):
##    configuracoes.gv_ds_pasta_base_serie = getattr(configuracoes,'pasta_base_serie', 'NAO_EXISTE_DIRETORIO_PROTOCOLADO')

configuracoes.gv_ds_pasta_base_obrigacao = os.path.join(configuracoes.gv_ds_diretorio_entrada, 'LEVCV115')
if not os.path.isdir(configuracoes.gv_ds_pasta_base_obrigacao) and os.path.isdir(getattr(configuracoes,'pasta_base_obrigacao', 'NAO_EXISTE_DIRETORIO_LEVCV115')):
    configuracoes.gv_ds_pasta_base_obrigacao = getattr(configuracoes,'pasta_base_obrigacao', 'NAO_EXISTE_DIRETORIO_LEVCV115')

configuracoes.gv_ds_dir_base_obrigacoes = os.path.join(configuracoes.gv_ds_diretorio_entrada, 'LEVCV115')
if not os.path.isdir(configuracoes.gv_ds_dir_base_obrigacoes) and os.path.isdir(getattr(configuracoes,'dir_base_obrigacoes', 'NAO_EXISTE_DIRETORIO_LEVCV115')):
    configuracoes.gv_ds_dir_base_obrigacoes = getattr(configuracoes,'dir_base_obrigacoes', 'NAO_EXISTE_DIRETORIO_LEVCV115')

log('...<<diretorio_entrada>>......' + str(configuracoes.gv_ds_diretorio_entrada))
log('...<<diretorio_saida>>........' + str(configuracoes.gv_ds_diretorio_saida))
log('...<<diretorio_work>>.........' + str(configuracoes.gv_ds_diretorio_work))
log('...<<diretorio_ready>>........' + str(configuracoes.gv_ds_diretorio_ready))
log('...<<pasta_base_serie>>.......' + str(configuracoes.gv_ds_pasta_base_serie))
log('...<<pasta_base_obrigacao>>...' + str(configuracoes.gv_ds_pasta_base_obrigacao))
log('...<<base_obrigacoes>>........' + str(configuracoes.gv_ds_dir_base_obrigacoes))

#- [PTITES-1639]

def leDadosArquivosControle() :
    dir_base_obrigacoes = configuracoes.gv_ds_dir_base_obrigacoes # - [PTITES-1639] #  configuracoes.dir_base_obrigacoes
    ano      = configuracoes.ano
    mes      = configuracoes.mes
    filial   = configuracoes.filial
    id_serie = configuracoes.id_serie
    uf       = configuracoes.uf

    dir_serie = os.path.join(dir_base_obrigacoes, uf, ano[2:], mes, 'TBRA', filial, 'SERIE', id_serie, 'PROTOCOLADO')

    log('- Diretorio da serie ..:', dir_serie)
    if not os.path.isdir(dir_serie) :
        log('Erro : Diretorio da serie invalido !')
        return False

    log('Iniciando processamento !')
    log('Buscando arquivos de controle.')
    lst_arqs = os.listdir(dir_serie)
    lst_arqs.sort()
    posC = 10 if int(ano) < 2017 else 28
    layout_controle = 'LayoutControleV3_Antigo' if int(ano) < 2017 else 'LayoutControleV3'

    # lst_layouts.append( [ 'controle', 'LayoutControleV3.csv' ] )
    # lst_layouts.append( [ 'controleAntigo', 'LayoutControleV3_Antigo.csv' ] )
    dados_volumes = {}

    configuracoes.dadosVolumes = dados_volumes

    log('- Layout utilizado :', layout_controle)
    qtd_arqs_controle = 0
    for arq in lst_arqs :
        if len(arq) >= posC and arq[posC] == 'C' :
            log('-'*100)
            qtd_arqs_controle += 1
            log('- Encontrado o arquivo de controle ..:', arq)
            volume = int(arq[-3:])
            log('> Volume :', arq[-3:])
            dados_volumes[volume] = {}
            path_arq = os.path.join( dir_serie, arq )
            encoding = comum.encodingDoArquivo( path_arq )
            fd = open(path_arq, 'r', encoding=encoding)
            reg_controle = fd.readline()
            fd.close()
            registro = layout.quebraRegistro(reg_controle, dic_layouts[layout_controle])
            idx_numPriDocMestre = dic_layouts[layout_controle]['dic_campos']['NUMEROPRIDOC']-1 
            idx_numUltDocMestre = dic_layouts[layout_controle]['dic_campos']['NUMEROULTDOC']-1
            dados_volumes[volume]['NUMEROPRIDOC'] = registro[idx_numPriDocMestre]
            log('> Numero do primeiro documento mestre ..: X%sX'% dados_volumes[volume]['NUMEROPRIDOC'].strip())
            if not dados_volumes[volume]['NUMEROPRIDOC'].strip() :
                log('Erro - Arquivo de controle inconsistente, falta valor do NumeroPrimeiroDocumento ...')
                return False
            
            dados_volumes[volume]['NUMEROULTDOC'] = registro[idx_numUltDocMestre]
            log('> Numero do ultimo documento mestre ....:', dados_volumes[volume]['NUMEROULTDOC'])
            if not dados_volumes[volume]['NUMEROULTDOC'].strip() :
                log('Erro - Arquivo de controle inconsistente, falta valor do NumeroUltimoDocumento ...')
                return False

            ##### Dados itens (I)
            idx_numPriDocItens = dic_layouts[layout_controle]['dic_campos']['NUMEROPRIDOCITENS'] -1
            idx_numUltDocItens = dic_layouts[layout_controle]['dic_campos']['NUMEROULTDOCITENS'] -1
            dados_volumes[volume]['NUMEROPRIDOCITENS'] = registro[idx_numPriDocItens]
            log('> Numero do primeiro documento itens ...:', dados_volumes[volume]['NUMEROPRIDOCITENS'])
            dados_volumes[volume]['NUMEROULTDOCITENS'] = registro[idx_numUltDocItens]
            log('> Numero do ultimo documento itens .....:', dados_volumes[volume]['NUMEROULTDOCITENS'])
    
    if qtd_arqs_controle == 0 :
        log('Erro - Nao foram encontrados os arquivos de Controle .... Verifique ...')
        return False

    log('-'*100)
    log('Dados dos arquivos de controle extraidos !')
    return True

def gerarNovosArquivos() :
    dir_base_obrigacoes = configuracoes.gv_ds_dir_base_obrigacoes # - [PTITES-1639] #  configuracoes.dir_base_obrigacoes
    ano      = configuracoes.ano
    mes      = configuracoes.mes
    filial   = configuracoes.filial
    id_serie = configuracoes.id_serie
    uf       = configuracoes.uf

    dir_obrigacoes = os.path.join(dir_base_obrigacoes, uf, ano[2:], mes, 'TBRA', filial, 'SERIE', id_serie, 'OBRIGACAO')
    log('- Diretorio de obrigacoes .:', dir_obrigacoes)
    if not os.path.isdir(dir_obrigacoes) :
        log('Erro - Diretorio de obrigacoes nao existe !')
        return False

    dt = datetime.datetime.now().strftime('%Y%m%d')
    dir_bkp_obrigacoes_raiz = os.path.join(dir_base_obrigacoes, uf, ano[2:], mes, 'TBRA', filial, 'SERIE', id_serie, 'bkp_%s'%(name_script) )
    dir_bkp_obrigacoes = os.path.join(dir_base_obrigacoes, uf, ano[2:], mes, 'TBRA', filial, 'SERIE', id_serie, 'bkp_%s'%(name_script), dt )
    # Inicio [PTITES-1639]
    dir_ready = os.path.join( configuracoes.gv_ds_diretorio_ready, id_serie )
    dir_work = os.path.join(  configuracoes.gv_ds_diretorio_work, id_serie )
    # Fim [PTITES-1639]
    pos = 10 if int(ano) < 2017 else 28

    if os.path.isdir( dir_bkp_obrigacoes ) :
        log('Limpando o diretorio de backup dos arquivos de obrigacao')
        lst_arqs = os.listdir(dir_bkp_obrigacoes)
        for item in lst_arqs :
            if os.path.isfile(os.path.join(dir_bkp_obrigacoes, item)) :
                os.remove( os.path.join(dir_bkp_obrigacoes, item) )
        os.rmdir(dir_bkp_obrigacoes)
        log(' - Backup limpo !')

    if not os.path.isdir(dir_ready) :
        os.makedirs(dir_ready)
        os.chmod(dir_ready, 0o777)
        
    else : 
        lst_arqs = os.listdir(dir_ready)
        for item in lst_arqs :
            path_item = os.path.join( dir_ready, item )
            if os.path.isfile(path_item) :
                #ALT001 - Inicio
                if  os.path.isfile(os.path.join(dir_bkp_obrigacoes, item)):
                #ALT001 - Fim
                    os.remove( os.remove( os.path.join(dir_bkp_obrigacoes, item) ) )

    if not os.path.isdir(dir_work) :
        os.makedirs(dir_work)
        os.chmod(dir_work, 0o777)
    else : 
        lst_arqs = os.listdir(dir_work)
        for item in lst_arqs :
            path_item = os.path.join( dir_work, item )
            if os.path.isfile(path_item) :
                #ALT001 - Inicio
                if  os.path.isfile(os.path.join(dir_bkp_obrigacoes, item)):
                #ALT001 - Fim
                    os.remove( os.remove( os.path.join(dir_bkp_obrigacoes, item) ) )

    if not os.path.isdir( dir_bkp_obrigacoes ) :
        log('Iniciando o backup dos arquivos de obrigacao')
        lst_arqs = os.listdir(dir_obrigacoes)
        os.makedirs(dir_bkp_obrigacoes)
        os.chmod(dir_bkp_obrigacoes, 0o777)
        
        try:
           os.chmod(dir_bkp_obrigacoes_raiz, 0o777)
        except Exception as e:
           log('Nao possivel dar permissao na raiz...')
           
        for item in lst_arqs :
            if os.path.isfile(os.path.join(dir_obrigacoes, item)) :
                if not item.endswith('.bad') and item[pos] != 'C' :
                    shutil.copy(os.path.join(dir_obrigacoes, item), dir_ready )
                    log("- Garantindo Permissao...")
                    os.chmod(os.path.join(dir_ready, item), 0o777)
        log(' - Backup criado !')

    log('Iniciando geracao dos novos arquivos!')
    layoutMestre = 'LayoutMestre' if int(ano) >= 2017 else 'LayoutMestre_Antigo'
    layoutItem = 'LayoutItem' if int(ano) >= 2017 else 'LayoutItem_Antigo'
    layoutCadastro = 'LayoutCadastro' if int(ano) >= 2017 else 'LayoutCadastro_Antigo'

    pos = 10 if int(ano) < 2017 else 28

    lst_arqs = os.listdir(dir_ready)
    lst_arqs.sort()
    dadosVolumes = configuracoes.dadosVolumes
    qtde_linhas = 0
    vol_anterior = 1
    ref_item = 0
    linhas = []
    linhasCad = []
    a_deletar = []
    tipo_anterior = False
    processarCadastro = False
    dic_ref_itens = {}
    for arq in lst_arqs :
        if processarCadastro :
            arq_cad.close()
        processarCadastro = False
        if arq[pos] == 'M' :
            v_layout = layoutMestre
            coluna_doc = 'NUMERO_NF'
            col_doc_controle = [ 'NUMEROPRIDOC', 'NUMEROULTDOC' ]
            processarCadastro = True if int(ano) < 2017 else False
        elif arq[pos] == 'I' :
            v_layout = layoutItem
            coluna_doc = 'NUMERO_NF'
            col_doc_controle = [ 'NUMEROPRIDOCITENS', 'NUMEROULTDOCITENS' ]
        elif arq[pos] == 'D' :
            v_layout = layoutCadastro
            coluna_doc = 'NUMERO_NF'
            col_doc_controle = [ 'NUMEROPRIDOC', 'NUMEROULTDOC' ]
            if int(ano) < 2017 :
                continue
        else :
            continue
        if arq[pos] != tipo_anterior :
            #### limpar processados ...
            for item in a_deletar :
                os.remove(item)
            a_deletar.clear()
            tipo_anterior = arq[pos]
        a_deletar.append(os.path.join(dir_ready,arq))

        idx_colunaDoc = dic_layouts[v_layout]['dic_campos'][coluna_doc]-1
        
        log('-'*100)
        encode = comum.encodingDoArquivo(os.path.join(dir_ready, arq))
        log('Processando arquivo ..:', arq, 'Encoding :', encode)
        arq_ready = open(os.path.join(dir_ready, arq), 'r', encoding=encode)
        nome_arquivo = arq.split('.')[0]
        linhas_arq = 0
        qtdeLinCad = 0
        if arq[pos] == 'M' and processarCadastro :
            nome_arq_cad = False
            for x in lst_arqs :
                if x[pos] == 'D' :
                    if x.split('.')[-1] == arq.split('.')[-1] :
                        nome_arq_cad = x
            if nome_arq_cad :
                encodeCad = comum.encodingDoArquivo(os.path.join(dir_ready, nome_arq_cad))
                log('- Processando junto o arquivo de cadastro :', nome_arq_cad)
                arq_cad = open(os.path.join(dir_ready, nome_arq_cad), 'r', encoding=encodeCad)
                a_deletar.append(os.path.join(dir_ready, nome_arq_cad))
            else :
                processarCadastro = False
            
        ### Pega o tamanho maximo do registro para o v_layout usado.
        keys_dic_registros_layout = [ x for x in dic_layouts[v_layout].keys() ]
        tam_max_registro = dic_layouts[v_layout][keys_dic_registros_layout[-1]][3] + 1

        for item_linha in arq_ready.readlines() :
            try :
                linha = item_linha.encode(encode).decode('utf-8') # if encode == 'iso-8859-1' else item_linha
            except :
                linha = item_linha

            if len(linha) < tam_max_registro :
                registro_temp = layout.quebraRegistro(linha.encode('utf-8'), v_layout)
                # log(dir(registro_temp[0]))
                registro = [ x.decode() for x in registro_temp ]
                #### Compara todas as colunas dos dois registros e verifica qual o campo ficou menor que o layout.
                #### Encontrado o campo, o mesmo eh acrescido de 1 espaco em branco em seu final.
                registro_errado = layout.quebraRegistro(linha, v_layout)
                for c in range(len(registro)) :
                    if len(registro[c]) < len(registro_errado[c]) :
                        #### Achei a coluna errada ... 
                        field, t, i, f = dic_layouts[v_layout][c+1]
                        log(registro)
                        log(field, t, i, f)
                        #### Alterando linha para ajustar o tamanho da coluna.
                        linha_temp = linha[:f] + ' ' + linha[f:]
                        linha = linha_temp[:]
            else :
                registro = layout.quebraRegistro(linha, v_layout)
            qtde_linhas += 1
            linhas_arq += 1
            ref_item += 1
            doc = int(registro[idx_colunaDoc])
            volume = False
            # log(doc)
            for vol in dadosVolumes.keys() :
                if doc >= int(dadosVolumes[vol][col_doc_controle[0]]) :
                    if doc <= int(dadosVolumes[vol][col_doc_controle[1]]) :
                        volume = vol
            if not volume :
                log('Erro - Nao encontrado o volume para o documento :', doc)
                log('Linha :', linhas_arq)
                log('Linha Quebrada:', registro )
                log(linha)
                return False
            
            if volume != vol_anterior :                
                log('#### Trocou de volume ... primeiro documento :', doc)
                log('REF_ITEM', ref_item)
                novo_arquivo = nome_arquivo + '.' + str(vol_anterior).rjust(3,'0')
                path_arq_work = os.path.join(dir_work, novo_arquivo)
                escreverArquivo(path_arq_work, linhas)
                linhas.clear()
                if processarCadastro :
                    novo_arquivo_cad = nome_arq_cad.split('.')[0] + '.' + str(vol_anterior).rjust(3,'0')
                    escreverArquivo(os.path.join(dir_work, novo_arquivo_cad),linhasCad)
                    linhasCad.clear()
                vol_anterior = volume
                ref_item = 1
                log('REF_ITEM', ref_item)
            
            if arq[pos] == 'I' :
                dic_ref_itens[volume] = dic_ref_itens.get(volume, {})
                if doc not in dic_ref_itens[volume].keys() :
                    dic_ref_itens[volume][doc] = ref_item 
            elif arq[pos] == 'M' :
                if volume not in dic_ref_itens.keys() :
                    log('- Montando dicionario de itens ... volume', volume)
                    #### Busca arquivo de Itens no diretorio WORK
                    nome_arq_itens = False
                    for arq_itens in os.listdir(dir_work) :
                        if arq_itens[pos] == 'I' and  int(arq_itens.split('.')[-1]) == volume :
                            nome_arq_itens = os.path.join(dir_work, arq_itens)
                    
                    if not nome_arq_itens :
                        #### Se nao achar busca no diretorio de obrigacoes.
                        for arq_itens in os.listdir(dir_obrigacoes) :
                            if arq_itens[pos] == 'I' and  int(arq_itens.split('.')[-1]) == volume :
                                nome_arq_itens = os.path.join(dir_obrigacoes, arq_itens)
                    if nome_arq_itens :
                        log('> Utilizando arquivo .:',nome_arq_itens)
                        dic_ref_itens[volume] = dic_ref_itens.get(volume, {})
                        encode_nome_arq_itens = comum.encodingDoArquivo(nome_arq_itens)
                        fd_arq_itens = open( nome_arq_itens, 'r', encoding = encode_nome_arq_itens )
                        y = 1
                        l = fd_arq_itens.readline() 
                        while l :
                            r = layout.quebraRegistro(l, layoutItem)
                            
                            idxcol = dic_layouts[layoutItem]['dic_campos'][coluna_doc]-1
                          
                            d = int(r[idxcol])
                            if d not in dic_ref_itens[volume].keys() :
                                dic_ref_itens[volume][d] = y
                            l = fd_arq_itens.readline() 
                            y += 1
                        fd_arq_itens.close()
                    log('- Dicionario de itens, volume', volume, 'criado!')
                    
                ##### Altera o campo REF_ITEM_NF
                field, t, i, f = dic_layouts[v_layout][dic_layouts[v_layout]['dic_campos']['REF_ITEM_NF']]
                ref_item_nf = dic_ref_itens[volume][doc]
                linha_temp = linha[:i-1]
                linha_temp += str(ref_item_nf).rjust(t,'0')
                linha_temp += linha[f:]
                if linhas_arq < 5 :
                    log('Alterado para ', ref_item_nf)
                    log(linha_temp)
                ##### Altera o HASH MD5 do registro 
                field, t, i, f = dic_layouts[v_layout][dic_layouts[v_layout]['dic_campos']['HASH_CODE_ARQ']]
                md5 = hashlib.md5()
                md5.update(linha_temp[:i-1].encode(encode))
                linha_temp = linha_temp[:i-1] + md5.hexdigest() + '\r\n'
                linha = linha_temp

            linhas.append(linha[:])
            if processarCadastro :
                qtdeLinCad += 1
                linhaCad = arq_cad.readline()
                linhasCad.append(linhaCad[:])
            
            
            if ( linhas_arq % 500000 ) == 0 :
                log('> Processadas', qtde_linhas, 'linhas.')
            
        #### Grava os ultimos registros ....
        novo_arquivo = nome_arquivo + '.' + str(vol_anterior).rjust(3,'0')
        path_arq_work = os.path.join(dir_work, novo_arquivo)
        escreverArquivo(path_arq_work, linhas)
        linhas.clear()
        
        if processarCadastro :
            novo_arquivo_cad = nome_arq_cad.split('.')[0] + '.' + str(vol_anterior).rjust(3,'0')
            escreverArquivo(os.path.join(dir_work, novo_arquivo_cad),linhasCad)
            linhasCad.clear()

        log('Total de linhas do arquivo ......:', linhas_arq)
        log('Total de linhas do arquivo CAD ..:', qtdeLinCad)
        arq_ready.close()
        del arq_ready
        if processarCadastro :
            arq_cad.close()

        log('- Arquivo processado com sucesso!')

    log(" R E S U M O ".center(50,'*'))
    log('Qtde de linhas processadas ...:', qtde_linhas)
    fechaArquivos()
    log('- Apagando arquivos ja processados.')
    for item in a_deletar :
        os.remove(item)
    
    log('- Excluindo arquivos do diretorio de OBRIGACOES ..')
    for item in os.listdir(dir_obrigacoes) :
        path_item = os.path.join(dir_obrigacoes, item)
        if os.path.isfile(path_item) :
            os.remove(path_item)

    log('- Movendo arquivos prontos.')
    log('- DIRETÓRIO DESTINO: ' , dir_obrigacoes)

    for item in os.listdir(dir_work) :
        log('- MOVENDO AQUIVO: ' , item)
        shutil.move(os.path.join(dir_work, item), dir_obrigacoes)

    return True


def escreverArquivo( arquivo, linhas ) :
    if arquivo not in dic_fd.keys() :
        dic_fd[arquivo] = open(arquivo, 'w', encoding = 'iso-8859-1')
    
    fd = dic_fd[arquivo]
    for linha in linhas :
        if linha.endswith('\r\n') :
            fd.write(linha)
        else :
            fd.write(linha.replace('\n','\r\n'))

    return True


def escreve( arquivo, linha ) :
    if arquivo not in dic_fd.keys() :
        flag = 'a' if os.path.isfile(arquivo) else 'w'
        dic_fd[arquivo] = open(arquivo, flag, encoding = 'iso-8859-1')
    fd = dic_fd[arquivo]
    fd.write(linha)

    return True

def fechaArquivos() :
    log('- Fechando arquivos prontos.')
    for key in dic_fd.keys() :
        dic_fd[key].close()
    return True


if __name__ == "__main__":
    ret = 0
    if len(sys.argv) > 1 :
        id_serie = sys.argv[1]
        configuracoes.id_serie = id_serie
    else :
        log('ERRO - Falta o parametro de execucao ID_SERIE.')
        log('- Exemplo :')
        log('        ./%s.py 1234455 '%(name_script))
        ret = 1
    if ret == 0 and not layout.carregaLayout() :
        ret = 2
    if ret == 0 and not comum.carregaConfiguracoes(configuracoes) :
        ret = 3
    # if ret == 0 and not comum.buscaDadosSerie(id_serie) :
    #     ret = 4
    variaveis = comum.buscaDadosSerie(configuracoes.id_serie)
    if not variaveis :
        ret = 4
    for var in variaveis:
        setattr(configuracoes, var, variaveis[var])

    if ret == 0 and not leDadosArquivosControle() :
        ret = 5
    if ret == 0 and not gerarNovosArquivos() :
        ret = 6
    log('-'*150)
    log('FIM da execucao!')
    status = 'SUCESSO' if ret == 0 else 'ERRO'
    log('STATUS da execucao :', status)
    sys.exit(ret)

