#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: GF
  MODULO ...: 
  SCRIPT ...: atualizaControle115.py
  CRIACAO ..: 12/01/2020
  AUTOR ....: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO : Atualizar os dados da tabela openrisow.CTR_IDENT_CNV115, de acordo com os 
              arquivos gerados.
              
----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 12/01/2020 - Welber Pena de Sousa - Kyros Tecnologia
            - Criacao do script.
----------------------------------------------------------------------------------------------
"""

import os
import sys
import cx_Oracle
# import unicodedata
# import fnmatch
import shutil
import datetime
import calendar

name_script = os.path.basename(__file__).split('.')[0]

dic_registros = {}
dic_layouts = {}
dic_campos = {}
variaveis = {}
dic_fd = {}

def print( *args ) :
    dt = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S |')
    if type(args[0]) == str and args[0].upper().startswith('ERRO') :
        tam = 15
        for i in args :
            tam += len(str(i))
        tam = 60 if tam < 60 else tam
        __builtins__.print(dt, "  E R R O  ".center(tam,'='))
        __builtins__.print(dt, '###', *args)
        __builtins__.print(dt, '='*tam )
    elif type(args[0]) == str and len(args[0]) > 0 and args[0][0:2] in ['- ', '> '] :
        t = '    '
        if args[0][0] == '>' :
            t += '    '
        __builtins__.print(dt, t, *args)
    else :
        __builtins__.print(dt, *args)


def buscaDadosSerie(id_serie) :
    print("Identificando mes e ano da serie .:", id_serie)
    banco = variaveis['banco']
    conexao = cx_Oracle.connect("gfcarga/vivo2019@"+banco, threaded=True)
    conexao.autocommit = False
    
    # try:
    if True :
        cursor = conexao.cursor()
        # try:
        if True :
            cursor.execute(""" 
                SELECT TO_CHAR(l.MES_ANO, 'YYYY') ANO, 
                       TO_CHAR(l.MES_ANO, 'mm') MES, 
                       l.SERIE,
                       l.FILI_COD
                FROM TSH_SERIE_LEVANTAMENTO l
                        INNER JOIN OPENRISOW.FILIAL f
                           ON l.EMPS_COD = f.EMPS_COD
                          AND l.FILI_COD = f.FILI_COD
                WHERE l.ID_SERIE_LEVANTAMENTO = :ID_SERIE """,
                (id_serie,)
            )
            linha = cursor.fetchone() 
            # linha = [ '2015', '01', '1', '0001' ]
            if linha :
                variaveis['ano']= linha[0]
                print('- Ano da serie ........:', linha[0])
                variaveis['mes']= linha[1]
                print('- Mes da serie ........:', linha[1])
                variaveis['serie']= linha[2].replace(' ','')
                variaveis['serie_original']= linha[2]
                print('- Serie ...............:', linha[2])
                variaveis['filial']= linha[3]
                print('- Filial ..............:', linha[3])
            else :
                print('Erro no cursor ao buscar referencias de Ano e Mes para o id_serie', id_serie)
                return False
        # except :
        #     print('Erro na conexao ao buscar referencias de Ano e Mes para o id_serie', id_serie)
        #     return False
        # finally:
            cursor.close()
    # except :
    #     print('Erro na funcao ao buscar referencias de Ano e Mes para o id_serie', id_serie)
    #     return False
    # finally:
        conexao.close()

    return True


def carregaLayout() :
    print('Carregando Layouts ....')
    lst_layouts = []
    lst_layouts.append( [ 'controle', 'LayoutControleV3.csv' ] )
    lst_layouts.append( [ 'controleAntigo', 'LayoutControleV3_Antigo.csv' ] )
    
    lst_layouts.append( [ 'mestre', 'LayoutMestre.csv' ] )
    lst_layouts.append( [ 'mestreAntigo', 'LayoutMestre_Antigo.csv' ] )
    
    lst_layouts.append( [ 'item', 'LayoutItem.csv' ] )
    lst_layouts.append( [ 'itemAntigo', 'LayoutItem_Antigo.csv' ] )

    lst_layouts.append( [ 'cadastro', 'LayoutCadastro.csv' ] )
    lst_layouts.append( [ 'cadastroAntigo', 'LayoutCadastro_Antigo.csv' ] )

    for layout, arq_layout in lst_layouts :
        if not os.path.isfile(os.path.join('..', 'unificado', 'layout', arq_layout)) :
            print('ERRO - Arquivo de Layout nao existe ... < %s >'%arq_layout)
            print('Diretorio de layouts :',os.path.join('..', 'unificado', 'layout',) )
            return False
        print('- Carregando layout : ', layout)
        fd = open(os.path.join('..', 'unificado', 'layout', arq_layout), 'r')

        dic_registros[layout] = {}
        dic_campos[layout] = {}
        for item in fd.readlines() :
            separador = ';' if item.__contains__(';') else ','
            linha = item.replace('\n','').split(separador)
            if linha[0].isdigit() :
                dic_registros[layout][int(linha[0])] = []
                for r in linha[1:] :
                    if r.isdigit() :
                        dic_registros[layout][int(linha[0])].append(int(r))
                    else :
                        dic_registros[layout][int(linha[0])].append(r)

                dic_campos[layout][linha[1]] = int(linha[0])
                # print(linha[0], '=',dic_registros[layout][int(linha[0])])
                # print(linha[1],'=',linha[0])

    print('Dicionario de Layouts criado !')        
    return True


def carregaConfiguracoes() :
    arq_cfg = '%s.cfg'%(name_script)
    print('Carregando dados do arquivo de configuracao .:', arq_cfg)
    if not os.path.isfile(os.path.join('.', arq_cfg)) :
        print('ERRO - Arquivo de configuracoes nao existe ... < %s >'%arq_cfg)
        return False
    fd = open(os.path.join('.', arq_cfg), 'r')
    for item in fd.readlines() :
        if not item.startswith('#') :
            linha = item.replace('\r\n','').replace('\n','').split('=')
            if len(linha) >= 2 :
                variaveis[linha[0]] = linha[1]
                print('- Setada a variavel', linha[0], '=',linha[1])
        
    return True


def encodingDoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r', encoding='iso-8859-1')
        t = fd.read()
        fd.close()
    except :
        return 'utf-8'

    return 'iso-8859-1'


def quebraRegistro(reg, layout) :
    itens_registro = []
    colunas = []
    for y in range( 1, len(dic_registros[layout].keys())+1 ) :
        field, t, i, f = dic_registros[layout][y]
        itens_registro.append( reg[i-1:f] )
        colunas.append(field)
        # print(field, '=', reg[i-1:f])
    # print(colunas)
    return itens_registro


def atualizaDados() :
    banco = variaveis['banco']
    print('Conectando ao banco de dados ...', banco)
    conexao = cx_Oracle.connect("gfcarga/vivo2019@"+banco, threaded=True)
    conexao.autocommit = False
    cursor = conexao.cursor()
    
    ano = variaveis['ano']
    mes = variaveis['mes']
    filial = variaveis['filial']
    id_serie = variaveis['id_serie']
    serie = variaveis['serie']
    
    print( 'Excluindo dados da tabela openrisow.CTR_IDENT_CNV115' )
    cmd_sql = """ 
    DELETE FROM  openrisow.CTR_IDENT_CNV115
    WHERE ctr_serie = '%s'
            and fili_cod = %s
            and ctr_apur_dtini = to_date( '01/%s/%s', 'dd/mm/yyyy' )
            and ctr_modelo in (21, 22)
            and CTR_IND_RETIF = 'S'
    """%( serie, filial, mes, ano )
    
    print(cmd_sql)
    cursor.execute(cmd_sql)
    
    ##### 2 - Busca no Diretório /portaloptrib/LEVCV115/SP/[ano]/[mes]/TBRA/[filial]/SERIE/[id serie]/OBRIGACAO 
    # o arquivo de controle de cada volume.
    pos = 10 if int(ano) < 2017 else 28

    dir_base_obrigacoes = variaveis.get('pasta_base_obrigacao', '/portaloptrib/LEVCV115/SP')
    dir_obrigacoes = os.path.join(dir_base_obrigacoes, ano[2:], mes, 'TBRA', filial, 'SERIE', id_serie, 'OBRIGACAO')
    # dir_obrigacoes = os.path.join(dir_base_obrigacoes, ano[2:], mes, 'TBRA', filial, 'SERIE', id_serie, 'OBRIGACAO', 'bkp_redistribuiVolumes', '20200110')
    print('- Diretorio de obrigacoes .:', dir_obrigacoes)
    if not os.path.isdir(dir_obrigacoes) :
        print('Erro - Diretorio de obrigacoes nao existe !')
        return False
    
    lst_arqs = os.listdir(dir_obrigacoes)
    arqs_processar = []

    for item in lst_arqs :
        if os.path.isfile(os.path.join(dir_obrigacoes, item)) :
            if not item.endswith('.bad') and item[pos] == 'C' :
                arqs_processar.append(item)
    
    arqs_processar.sort()

    #### Preparando dados para o insert.
    dic_field_insert = {}

    ### Dicionario de campos do insert
    # Campo = [ Tipo de Dado, 
    #           Contedudo a ser incluido (2017 a atualmente) , 
    #           Contedudo a ser incluido (2011 a 2016)  
    #       ]

    ultDiaDoMes = calendar.monthrange( int(ano), int(mes))[1]

    dic_field_insert['EMPS_COD'] = ['fixo', 'TBRA', 'fixo', 'TBRA']
    dic_field_insert['FILI_COD'] = [ 'fixo', filial, 'fixo', filial ]
    dic_field_insert['FILI_COD_IE'] = [ 'VARCHAR', 'IE', 'VARCHAR', 'IE' ]
    dic_field_insert['CTR_APUR_DTINI'] = [ 'numero', "to_date( '01/%s/%s', 'dd/mm/yyyy' )"%(mes,ano), 'numero',  "to_date( '01/%s/%s', 'dd/mm/yyyy' )"%(mes,ano) ]
    dic_field_insert['CTR_APUR_DTFIN'] = [ 'numero', "to_date( '%s/%s/%s', 'dd/mm/yyyy' )"%(ultDiaDoMes, mes, ano), 'numero', "to_date( '%s/%s/%s', 'dd/mm/yyyy' )"%(ultDiaDoMes, mes, ano)]
    dic_field_insert['CTR_SERIE'] = [ 'VARCHAR_SEM_ESPACO', 'Serie', 'fixo', serie ]
    dic_field_insert['CTR_MODELO'] = [ 'VARCHAR', 'Modelo', 'fixo', '22' ]
    dic_field_insert['CTR_VOLUME'] = [ 'VARCHAR', 'Volume', 'variavel', 'volume' ]
    dic_field_insert['CTR_IND_RETIF'] = [ 'fixo', 'S', 'fixo', 'S' ]
    dic_field_insert['CTR_TIP_MIDIA'] = [ 'fixo', 'DVD-R', 'fixo', 'DVD-R' ]
    dic_field_insert['CTR_QTD_MESTRE'] = [ 'NUMBER', 'QtdeRegMestre', 'NUMBER', 'QtdeRegMestre' ]
    dic_field_insert['CTR_QTD_NFCANC'] = [ 'NUMBER', 'QtdeNFcanceladas', 'NUMBER', 'QtdeNFcanceladas' ]
    dic_field_insert['CTR_DTA_NFINI'] = [ 'DATE', 'DataEmissaoPriDoc', 'DATE', 'DataEmissaoPriDoc' ]
    dic_field_insert['CTR_DTA_NFFIN'] = [ 'DATE', 'DataEmissaoUltDoc', 'DATE', 'DataEmissaoUltDoc' ]
    dic_field_insert['CTR_NUM_NFINI'] = [ 'NUMBER', 'NumeroPriDoc', 'NUMBER', 'NumeroPriDoc' ]
    dic_field_insert['CTR_NUM_NFFIN'] = [ 'NUMBER', 'NumeroUltDoc', 'NUMBER', 'NumeroUltDoc' ]
    dic_field_insert['CTR_NF_VLRTOTAL'] = [ 'NUMBER_DIV_100', 'ValorTotal', 'NUMBER_DIV_100', 'ValorTotal' ]
    dic_field_insert['CTR_NF_VLRBASE'] = [ 'NUMBER_DIV_100', 'BCIcms', 'NUMBER_DIV_100', 'BCIcms' ]
    dic_field_insert['CTR_NF_VLRICMS'] = [ 'NUMBER_DIV_100', 'ICMS', 'NUMBER_DIV_100', 'ICMS' ]
    dic_field_insert['CTR_NF_VLRISEN'] = [ 'NUMBER_DIV_100', 'OpIsentas', 'NUMBER_DIV_100', 'OpIsentas' ]
    dic_field_insert['CTR_NF_VLROUTRAS'] = [ 'NUMBER_DIV_100', 'ValoresNaoBC', 'NUMBER_DIV_100', 'ValoresNaoBC' ]
    dic_field_insert['CTR_NF_NOMARQ'] = [ 'VARCHAR', 'NomeArqMestre', 'VARCHAR', 'NomeArqMestre' ]
    dic_field_insert['CTR_CODH_ARQNF'] = [ 'VARCHAR', 'CodAutenticacao', 'VARCHAR', 'CodAutenticacao' ]
    dic_field_insert['CTR_QTD_ITEM'] = [ 'NUMBER', 'QtdeRegItem', 'NUMBER', 'QtdeRegItem' ]
    dic_field_insert['CTR_QTD_ITEMCANC'] = [ 'NUMBER', 'QtdeItensCancelados', 'NUMBER', 'QtdeItensCancelados' ]
    dic_field_insert['CTR_DTA_ITEMINI'] = [ 'DATE', 'DataEmissaoPriDocItens', 'DATE', 'DataEmissaoPriDocItens' ]
    dic_field_insert['CTR_DTA_ITEMFIN'] = [ 'DATE', 'DataEmissaoUltDocItens', 'DATE', 'DataEmissaoUltDocItens' ]
    dic_field_insert['CTR_NUM_ITEMINI'] = [ 'NUMBER', 'NumeroPriDocItens', 'NUMBER', 'NumeroPriDocItens' ]
    dic_field_insert['CTR_NUM_ITEMFIN'] = [ 'NUMBER', 'NumeroUltDocItens', 'NUMBER', 'NumeroUltDocItens' ]
    dic_field_insert['CTR_ITEM_VLRTOTAL'] = [ 'NUMBER_DIV_100', 'Total', 'NUMBER_DIV_100', 'Total' ]
    dic_field_insert['CTR_ITEM_VLRDESC'] = [ 'NUMBER_DIV_100', 'Descontos', 'NUMBER_DIV_100', 'Descontos' ]
    dic_field_insert['CTR_ITEM_VLRDESP'] = [ 'NUMBER_DIV_100', 'Acrescimos', 'NUMBER_DIV_100', 'Acrescimos' ]
    dic_field_insert['CTR_ITEM_VLRBASE'] = [ 'NUMBER_DIV_100', 'BCIcmsTotal', 'NUMBER_DIV_100', 'BCIcmsTotal' ]
    dic_field_insert['CTR_ITEM_VLRICMS'] = [ 'NUMBER_DIV_100', 'ICMSTotal', 'NUMBER_DIV_100', 'ICMSTotal' ]
    dic_field_insert['CTR_ITEM_VLRISEN'] = [ 'NUMBER_DIV_100', 'OpIsentasTotal', 'NUMBER_DIV_100', 'OpIsentasTotal' ]
    dic_field_insert['CTR_ITEM_VLROUTR'] = [ 'NUMBER_DIV_100', 'ValoresNaoBCTotal', 'NUMBER_DIV_100', 'ValoresNaoBCTotal' ]
    dic_field_insert['CTR_ITEM_NOMARQ'] = [ 'VARCHAR', 'NomeArqItem', 'VARCHAR', 'NomeArqItem' ]
    dic_field_insert['CTR_CODH_ARQITEM'] = [ 'VARCHAR', 'CodAutenticacaoItem', 'VARCHAR', 'CodAutenticacaoItem' ]
    dic_field_insert['CTR_QTD_CLI'] = [ 'NUMBER', 'QtdeCadastroDest', 'NUMBER', 'QtdeCadastroDest' ]
    dic_field_insert['CTR_CLI_NOMARQ'] = [ 'VARCHAR', 'NomeArqCadastro', 'VARCHAR', 'NomeArqCadastro' ]
    dic_field_insert['CTR_CODH_ARQCLI'] = [ 'VARCHAR', 'CodAutenticacaoCadastro', 'VARCHAR', 'CodAutenticacaoCadastro' ]
    dic_field_insert['CTR_CODH_REG'] = [ 'VARCHAR', 'CodAutenticacaoRegistro', 'VARCHAR', 'CodAutenticacaoRegistro' ]
    dic_field_insert['CTR_SER_ORI'] = [ 'VARCHAR', 'Serie', 'fixo', variaveis['serie_original'] ]
    dic_field_insert['CTR_VAL_RED'] = [ 'NUMERO', 'null', 'NUMERO', 'null' ]
    dic_field_insert['CTR_DT_GER'] = [ 'NUMERO', 'SYSDATE', 'NUMERO', 'SYSDATE' ]
    dic_field_insert['CTR_USUA_GER'] = [ 'Fixo', 'TESHUVA_AJ_VOL', 'fixo', 'TESHUVA_AJ_VOL' ]

    ##### 3 – Para cada arquivo encontrado, realizar a insert na 
    # tabela openrisow.CTR_IDENT_CNV115, conforme detalhado
    print('Processando arquivos ...')
    layout_controle = 'controleAntigo' if int(ano) < 2017 else 'controle'
    tp  = 2 if int(ano) < 2017 else 0
    val = 3 if int(ano) < 2017 else 1
    for arq in arqs_processar :
        print('-'*100)
        print('Atualizando dados do arquivo .:', arq)
        volume = arq[-3:]
        print('> Volume :', arq[-3:])
        path_arq = os.path.join( dir_obrigacoes, arq )
        encoding = encodingDoArquivo( path_arq )
        fd = open(path_arq, 'r', encoding=encoding)
        reg_controle = fd.readline()
        fd.close()
        registro = quebraRegistro(reg_controle, layout_controle)
        # print(reg_controle)
        # print(registro)

        print('Gerando insert de dados.')
        fields = ""
        values = ""

        for field in dic_field_insert.keys() :
            fields += ", %s"%(field) if fields != "" else field
            dados = dic_field_insert[field]
            tipo = dados[tp]
            valor = dados[val]
            if values != "" :
                values += ", "
            if tipo.upper() == 'FIXO' :
                values += "'%s'"%(valor)
            elif tipo.upper() == 'NUMERO' :
                values += valor
            elif tipo.upper() == 'VARCHAR' :
                # print (registro)
                # print(dic_campos[layout_controle][valor]-1)
                # print(len(registro))
                # print(layout_controle)
                # print('--->', dic_registros[layout_controle].keys())
                values += "TRIM('%s')"%(registro[dic_campos[layout_controle][valor]-1])
            elif tipo.upper() == 'VARCHAR_SEM_ESPACO' :
                values += "'%s'"%(registro[dic_campos[layout_controle][valor]-1].replace(' ',''))
            elif tipo.upper() == 'NUMBER' :
                values += "%s"%(registro[dic_campos[layout_controle][valor]-1])
            elif tipo.upper() =='NUMBER_DIV_100' :
                values += "%s"%(int(registro[dic_campos[layout_controle][valor]-1])/100.00)
            elif tipo.upper() == 'DATE' :
                values += "to_date(%s,'yyyymmdd')"%(registro[dic_campos[layout_controle][valor]-1])
            elif tipo.upper() == 'VARIAVEL' :
                values += "'%s'"%(eval(valor))
            else :
                print('Erro - Falta valores para o tipo de dados', tipo)
                print('- Campo que esta sendo trabalhado .:', field)
                return False
        
        cmd_sql = """
        INSERT INTO openrisow.CTR_IDENT_CNV115
        ( %s )
        VALUES
        ( %s )
        """%(fields, values)
        print(cmd_sql)
        cursor.execute(cmd_sql)

    print('Realizando o COMMIT das operacoes.')
    conexao.commit()
    print('Fechando cursores e conexoes!')
    cursor.close()
    conexao.close()

    return True


if __name__ == "__main__":
    ret = 0
    print('INICIO da execucao do script ...: %s.py'%(name_script))
    print('='*150)
    if len(sys.argv) > 1 :
        id_serie = sys.argv[1]
        variaveis['id_serie'] = id_serie
    else :
        print('ERRO - Falta o parametro de execucao ID_SERIE.')
        print('- Exemplo :')
        print('        ./%s.py 1234455 '%(name_script))
        ret = 1
    if ret == 0 and not carregaLayout() :
        ret = 2
    if ret == 0 and not carregaConfiguracoes() :
        ret = 3
    if ret == 0 and not buscaDadosSerie(id_serie) :
        ret = 4
    if ret == 0 and not atualizaDados() :
        ret = 6
    print('-'*150)
    print('FIM da execucao!')
    status = 'SUCESSO' if ret == 0 else 'ERRO'
    print('STATUS da execucao :', status)
    sys.exit(ret)

