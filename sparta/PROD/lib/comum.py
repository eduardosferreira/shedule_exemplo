"""
----------------------------------------------------------------------------------------------
  BIBLIOTECA .: comum.py
  CRIACAO ....: 28/02/2020
  AUTOR ......: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO ..: Biblioteca de funcoes comuns mais utilizadas.

                - log( na tela ou em arquivo )
                - carregaConfiguracoes()
                - buscaDadosSerie()
                - encodingDoArquivo()
                - addParametro()
                - getParametro()
                - validarParametros()
              
----------------------------------------------------------------------------------------------
  HISTORICO ..: 
    * 28/02/2020 - Welber Pena de Sousa - Kyros Tecnologia
        - Criacao do script.
    
    * 02/02/2021 - Welber Pena de Sousa - Kyros Tecnologia
        - Alterada a função do log para gerar erros com separação de linha através do \n
    
    * 04/03/2021 - Welber Pena de Sousa - Kyros Tecnologia
        - Criadas as funções para tratar os parametros do script.
            - addParametro()
            - getParametro()
            - validarParametros()
            
    * 13/01/2022 - Eduardo da Silva Ferreira - Kyros Tecnologia
        PTITES-1367 : Acrescentar a informação da origem do protocolado
        https://jira.telefonica.com.br/browse/PTITES-1367
        https://wikicorp.telefonica.com.br/x/7K8PDQ
        https://wikicorp.telefonica.com.br/x/JKMPDQ
        Funcao alterada:
            - buscaDadosSerie()
----------------------------------------------------------------------------------------------
"""

import datetime
import cx_Oracle
import os
import sys
import atexit
import traceback
import configuracoes
import sql

parametros = {}
idx_parametros = {}


def log_close() :
    log('-'*150)
    log('FIM da execucao!')
    status = 'SUCESSO' if not log.ret else 'ERRO'
    log('STATUS da execucao :', status)
    if log.gerar_log_em_arquivo :
        log.gerar_log_em_arquivo.close()
        log.gerar_log_em_arquivo = False


def log( *args ) :
    texto = []
    
    if log.first :
        texto.append('INICIO da execucao do script ...: %s.py'%(os.path.basename(sys.argv[0]).split('.')[0]))
        texto.append('='*150)
        atexit.register(log_close)
        log.first = False

    #### Trata o texto a ser impresso no log.
    dt = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S |')
    if type(args[0]) == str and args[0].upper().startswith('ERRO') :
        log.ret += 1
        txt = ' '.join( str(x) if not isinstance(x, bytes) else x.decode() for x in args )
        max_tam = 0
        max_lin = 160
        for i in txt.split('\n') :
            tam = 15 + len(str(i))
            max_tam = tam if tam > max_tam else max_tam

        tam = 60 if max_tam < 60 else max_tam if max_tam < max_lin else max_lin
        texto.append('')
        texto.append('### ' + "  E R R O  ".center(tam,'='))
        
        for lin in txt.split('\n') :
            if len(lin) > max_lin :
                b = """"""
                for t in lin.split(' ') :
                    if b :
                        b += ' '
                    if len(b.split('\n')[-1]) + len(t) < max_lin :
                        b += t
                    else :
                        b += '\n '+ t
                for linb in b.split('\n') :
                    texto.append( '### %s'%(linb))
            else :
                texto.append( '### %s'%(lin))

        texto.append( '### ' + '='*tam )
        texto.append('')
            
        txt = traceback.format_exc()
        if not txt.startswith('NoneType: None') :
            tam_max_txt = 100
            if not isinstance(txt, list) :
                txt = txt.split('\n')
            for lin in txt :
                print(lin)
                if (len(lin.strip()) + 15) > tam_max_txt :
                    tam_max_txt = len(lin.strip()) + 20
                    if (tam_max_txt % 2) != 0 :
                        tam_max_txt += 1
                    print( '>>>>>',  tam_max_txt)

            texto.append('*-'*(int(tam_max_txt/2)))
            texto.append('  Ultima EXCECAO levantada pelo python  '.center(tam_max_txt,'='))
            texto.append('')
            
            for lin in txt :
                for l in lin.split('\n') :
                    if l :
                        texto.append('[TRACEBACK] > ' + l)
                        msg = l
            texto.append('')
            texto.append('='*tam_max_txt)
            texto.append('*-'*(int(tam_max_txt/2)))
            texto.append('')


    elif type(args[0]) == str and len(args[0]) > 0 and args[0][0:2] in ['- ', '> '] :
        txt = ' '.join( str(x) if not isinstance(x, bytes) else x.decode() for x in args )[2:]
        for lin in txt.split('\n') :
            t = '    '
            if args[0][0] == '>' :
                t += '    '
            texto.append( '%s%s %s'%( t, args[0][0], lin ) )
    else :
        txt = ' '.join( str(x) if not isinstance(x, bytes) else x.decode() for x in args )
        for lin in txt.split('\n') :
            texto.append(lin)
    
    #### Verifica se vai gerar arquivo de log.
    try :
        if log.gerar_log_em_arquivo :
            if isinstance( log.gerar_log_em_arquivo, (bool, str) ) :
                if isinstance( log.gerar_log_em_arquivo, str ) :
                    name_script = log.gerar_log_em_arquivo
                else :
                    name_script = os.path.basename(sys.argv[0]).replace('.py','')
                path_log = os.path.join(configuracoes.dir_log, datetime.datetime.now().strftime('%Y%m%d') )
                # print('DIR_LOG  >>>', path_log)
                # path_log = './log/%s'%( datetime.datetime.now().strftime('%Y%m%d') )
                log.path_log = path_log
                if not os.path.isdir(path_log) :
                    os.makedirs( path_log )
                log.gerar_log_em_arquivo = open( '%s/%s_%s.log'%(path_log, name_script, datetime.datetime.now().strftime('%Y%m%d%H%M%S')), 'w' )
                # atexit.register( log.gerar_log_em_arquivo.close )
                log('Diretorio de LOG do script .....: %s'%(log.path_log))
    except Exception as e :
        log.gerar_log_em_arquivo = False
        print( dt, '[Exception LOG] - %s'%(e) )

    for lin in texto :
        print( dt, lin )
        if log.gerar_log_em_arquivo :
            try :
                log.gerar_log_em_arquivo.write( dt + ' ' + lin + '\n' )
            except Exception as e :
                print( dt, '[Exception LOG] - %s'%(e) )
                try:
                    log.gerar_log_em_arquivo.write( dt + ' ' + '[Exception LOG] - %s'%(e) + '\n' )
                except :
                    pass


log.gerar_log_em_arquivo = False
log.first = True
log.ret = 0
__builtins__['log'] = log


def carregaConfiguracoes(configuracoes, arq_cfg = None) :
    """
    Se não for passado o arq_cfg para buscar as configurações a propria função 
    chama recursivamente e busca por um arquivo de configuracao com o nome igual o script . 
        - Primeiro procura pelo arquivo passado como parametro arq_cfg :
            - Caso encontrado, carrega as configuracoes do mesmo.
            - Casa NAO encontrado, retorna o valor False .
        
        - Segundo procura pelo arquivo de configuracao <nome do script>.cfg 
            - Caso encontrado, carrega as configuracoes gravadas no mesmo.
            - Caso NAO encontrado, retorna o valor False.
    """
    config = {}
    if not arq_cfg :
        arq_cfg = '%s.cfg'%(os.path.basename(sys.argv[0]).replace('.py',''))
        cfg2 = carregaConfiguracoes(configuracoes, arq_cfg) 
        if cfg2 :
            for k in cfg2.keys() :
                config[k] = cfg2[k]

    else :
        path_arq_cfg = os.path.join( '.', arq_cfg )
        
        if not os.path.isfile( path_arq_cfg ) :
            return False

        if arq_cfg :
            log('Carregando dados do arquivo de configuracao .:', arq_cfg)
            log('  - Path .:', path_arq_cfg)
            fd = open(path_arq_cfg, 'r')
            for item in fd.readlines() :
                if not item.startswith('#') :
                    if item.__contains__('=') :
                        linha = item.replace('\n','').split('=')
                        if len(linha) >= 2 :
                            if linha[0].lower().__contains__('dir') and linha[1].__contains__('/') :
                                setattr(configuracoes, linha[0].strip(), configuracoes.raiz + linha[1].strip().replace('/', configuracoes.SD))
                            else :
                                if len(linha) == 2 :
                                    setattr(configuracoes, linha[0].strip(), linha[1].strip())
                                else :
                                    setattr(configuracoes, linha[0].strip(), '='.join( x.strip() for x in linha[1:]))
                            
                            config[linha[0].strip()] = getattr(configuracoes, linha[0].strip(), False )
                            log('- Setada a configuracao', linha[0].strip(), '=', getattr(configuracoes, linha[0].strip(), False ) if not ( linha[0].lower().__contains__('pwd') or linha[0].lower().__contains__('pass') or linha[0].lower().__contains__('senha') ) else '*'*len(linha[1].strip()) )
            fd.close()
        else :
            return False

    return config if config else False


def addParametro(nomeParametro, identificador = None, descricao = '', obrigatorio = False, exemplo = None, default = False) :
    global parametros
    global idx_parametros
    idx = len(parametros.keys()) + 1
    parametros[idx] = {}
    param = parametros[idx]
    param['nome'] = nomeParametro
    param['identificador'] = identificador
    param['descricao'] = descricao
    param['obrigatorio'] = obrigatorio
    param['exemplo'] = exemplo
    param['default'] = default

    idx_parametros[nomeParametro] = idx
    return True


def getParametro(nomeParametro) :
    global parametros
    global idx_parametros
    if nomeParametro in idx_parametros.keys() :
        return parametros[idx_parametros[nomeParametro]].get('valor', parametros[idx_parametros[nomeParametro]]['default'] )
    return False


def validarParametros() :
    name_script = os.path.basename(sys.argv[0]).replace('.py', '')
    global parametros

    erro = False

    ### Valida a quantidade de parametros ...
    qt_parametros_obrigatorios = 1
    parametros_identificados = False
    for k in parametros.keys() :
        if parametros[k]['obrigatorio'] :
            qt_parametros_obrigatorios += 1
        if parametros[k]['identificador'] :
            parametros_identificados = True
        if parametros_identificados and parametros[k]['obrigatorio'] and parametros[k]['identificador'] not in sys.argv :
            erro = True

    if parametros_identificados :
        if len(sys.argv) < qt_parametros_obrigatorios :
            erro = True
    else :
        if len(sys.argv) < qt_parametros_obrigatorios :
            erro = True

    if not erro and len(sys.argv) > 1 :
        ### Recebe os valores dos parametros ...
        lst_params = sys.argv[1:]
        ident = False
        ident_ant = False
        
        for idx in range(len(lst_params)) :
            if not parametros_identificados : 
                ### Parametros sequenciais 
                # Exemplo :
                #    ./scriptTeste.py param1 param2 param3
                parametros[idx+1]['valor'] = lst_params[idx]
                # log(' - %s = %s'%( parametros[idx+1]['nome'], parametros[idx+1]['valor'] ))
            else :
                ### Parametros identificados
                # Exemplo :
                #    ./scriptTeste.py -I1 param1 -I2 param2 continuacao.param2 -I3 param3
                for k in parametros.keys() :
                    if parametros[k]['identificador'] == lst_params[idx] :
                        ident_ant = ident
                        ident = k
                
                # if ident_ant and ident_ant != ident :
                    # log(' - %s = %s'%( parametros[ident_ant]['nome'], parametros[ident_ant]['valor'] ))

                if ident :
                    if parametros[ident]['identificador'] != lst_params[idx] :
                        if parametros[ident].get('valor', False) :
                            parametros[ident]['valor'] += ' %s'%( lst_params[idx] )
                        else :
                            parametros[ident]['valor'] = lst_params[idx] 
                else :
                    erro = True

    
    txt = '     ./%s.py '%(name_script)
    max_tam_nome = 10
    max_tam_ident = 13
    max_tam_exemplo = 7
    for k in parametros.keys() :
        if parametros_identificados :
            txt += '%s %s '%( parametros[k]['identificador'], parametros[k]['exemplo'] if parametros[k]['exemplo'] else '< %s >'%( parametros[k]['nome'].upper() ) )
        else :
            txt += '%s '%( parametros[k]['exemplo'] if parametros[k]['exemplo'] else '< %s >'%( parametros[k]['nome'].upper() ) )
        max_tam_nome = max_tam_nome if len(parametros[k]['nome']) +3 < max_tam_nome else len(parametros[k]['nome']) +3
        if parametros[k]['exemplo'] :
            max_tam_exemplo = max_tam_exemplo if len(parametros[k]['exemplo']) +2 < max_tam_exemplo else len(parametros[k]['exemplo']) +2

    exemplo = txt
    
    
    txt = '\nNome'.ljust(max_tam_nome)
    txt += '  | '
    txt += 'Identificador'.ljust(max_tam_ident)
    txt += ' | '
    txt += 'Exemplo'.ljust(max_tam_exemplo)
    txt += ' | '
    txt += 'Obrigatorio'.ljust(11)
    txt += ' | '
    txt += 'Descricao \n'
    t = len(txt)+70 if len(txt)+50 < 190 else 190
    txt = ('-'*t) + txt 
    txt = 'Os parametros do script são :\n\n' + txt
    txt += '-'*t +'\n'

    for k in parametros.keys() :
        txt += '%s | %s | %s | %s | %s'%(   parametros[k]['nome'].ljust(max_tam_nome), 
                                            parametros[k]['identificador'].ljust(max_tam_ident) if parametros[k]['identificador'] else '-'.center(max_tam_ident, ' '),
                                            parametros[k].get('exemplo', '').ljust(max_tam_exemplo, ' ') if parametros[k].get('exemplo', False) else ' '.ljust(max_tam_exemplo, ' '), 
                                            'SIM'.ljust(11) if parametros[k]['obrigatorio'] else '-'.ljust(11) ,
                                            parametros[k]['descricao'] )
        txt += '\n'

    txt += 'Exemplo de execução :\n'
    txt += exemplo
    txt += '\n'
    txt += '-'*t
    txt += '\n'
    validarParametros.help = txt
    
    if erro :
        log('ERRO - Erro nos parametros passados para o script.')
        imprimeHelp()
        return False

    if parametros :
        log('Parametros passados para o script ..:')
        for k in parametros :
            if parametros[k].get('valor', False) :
                log(' > %s = %s'%( parametros[k]['nome'].upper(), parametros[k]['valor'] ))
        log('-'*150)

    return True


def imprimeHelp() :
    for lin in validarParametros.help.split('\n') :
        log(lin)
    return True


def buscaDadosSerie(id_serie) :
    """
    Essa funcao conecta no banco de dados citado em configuracoes.py
    Busca pelos dados da serie passada como parametro, na tabela TSH_SERIE_LEVANTAMENTO
    E retorna um dicionario na seguinte formacao :
      {
         ano       : 2020,
         mes       : '02',
         serie     : 'C',
         filial    : '',
         uf        : 'SP',
         empresa   : 'TBRA',
         dir_serie : '/portaloptrib/LEVCV115/SP/20/02/0001/TBRA/C/id_serie
      }
    """
    # print(dir())
    log("Identificando dados da serie ..:", id_serie)
    obj_sql = sql.geraCnxBD(configuracoes)
    variaveis = {}
    try:
        obj_sql.executa("""
            SELECT TO_CHAR(l.MES_ANO, 'YYYY') ANO,
                       TO_CHAR(l.MES_ANO, 'mm') MES,
                       l.SERIE,
                       l.FILI_COD,
                       l.EMPS_COD,
                       f.UNFE_SIG,
                       f.UNFE_SIG||'/'||TO_CHAR(l.MES_ANO, 'YY/MM')||'/'||l.EMPS_COD||'/'||l.FILI_COD||'/SERIE/'||l.ID_SERIE_LEVANTAMENTO,
                      TO_CHAR(l.MES_ANO, 'dd/mm/yyyy') data_ini,
                      INDICADOR_RETIFICACAO,
                      SEQUENCIA
                      , ORIGEM_PROTOCOLADO -- <<PTITES-1367>>
                FROM GFCARGA.TSH_SERIE_LEVANTAMENTO l
                        INNER JOIN OPENRISOW.FILIAL f
                           ON l.EMPS_COD = f.EMPS_COD
                          AND l.FILI_COD = f.FILI_COD
                WHERE l.ID_SERIE_LEVANTAMENTO = :ID_SERIE """,
            (id_serie,)
        )
        linha = obj_sql.fetchone() 
        # linha = [ '2015', '01', '1', '0001' ]
        variaveis['id_serie'] = id_serie
        if linha :
            variaveis['ano']= linha[0]
            log('- Ano da serie ..............:', variaveis['ano'])

            variaveis['mes']= linha[1]
            log('- Mes da serie ..............:', variaveis['mes'])

            variaveis['serie'] = linha[2].replace(' ','')
            variaveis['serie_original'] = linha[2]
            log('- Serie .....................:', variaveis['serie'])
            log('- Serie original ............:', variaveis['serie_original'])

            variaveis['filial'] = linha[3]
            log('- Filial ....................:', variaveis['filial'])

            variaveis['empresa'] = linha[4]
            log('- Empresa ...................:', variaveis['empresa'])

            variaveis['uf'] = linha[5]
            log('- UF ........................:', variaveis['uf'])

            variaveis['sub_dir_serie'] = linha[6]
            log('- Sub Diretório Série .......:', variaveis['sub_dir_serie'])

            ## ALT004 - Inicio
            variaveis['data_ini_apuracao'] = linha[7]
            log('- Data inicio apuracao ......:', variaveis['data_ini_apuracao'])
            ## ALT004 - Fim

            ## ALT005 - Inicio
            variaveis['indicador_retificacao'] = linha[8]
            log('- Indicador de Retificacao ..:', variaveis['indicador_retificacao'])
            
            variaveis['sequencia'] = linha[9]
            log('- Sequencia .................:', variaveis['sequencia'])
            ## ALT005 - Fim
            
            ## -- <<PTITES-1367>> - Inicio
            variaveis['origem_protocolado'] = linha[10]
            log('- Origem do protocolado .....:', variaveis['origem_protocolado'])
            ## -- <<PTITES-1367>> - Fim
            
            #ALT001 - Inicio
            if configuracoes.ambiente == 'DEV':
                variaveis['dir_serie'] = os.path.join(configuracoes.raiz, linha[6])
            else:
                variaveis['dir_serie'] = "/portaloptrib/LEVCV115/"+linha[6]
            
            log('- Diretorio da serie ........:', variaveis['dir_serie'])
            #ALT001 - Fim

            obj_sql.executa("""
                    SELECT u.modelo_nf
                    FROM gfcarga.TSH_TAB_TP_UTILIZ_CLASS_FIS u
                    WHERE u.SERIE = :SERIE""",
                    ( variaveis['serie_original'], )
                )
            linha = obj_sql.fetchone() 
            variaveis['modelo_nf'] = linha[0]
            log('- Modelo NF .................:', linha[0])

        else :
            log('Erro no cursor ao buscar referencias de Ano e Mes para o id_serie', id_serie)
            return False
    except Exception as e :
        print("AKIIIIII", e )
        log('Erro na funcao ao buscar referencias de Ano e Mes para o id_serie', id_serie)
        return False

    return variaveis


def encodingDoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r')
        b = fd.read()
    except :
        fd = open(path_arq, 'rb' )
        b = fd.read()
        
    fd.close()
    try :
        if type(b) == bytes :
            c = b.decode('iso-8859-1')
        else :
            c = b.encode('iso-8859-1')
    except :
        try :
            if type(b) == bytes :
                c = b.decode('latin-1')
            else :
                c = b.encode('latin-1')
            return 'latin-1'
        except :
            return 'utf-8'
    return 'iso-8859-1'


def ordenaListaDicionarios( lista_desordenada, lista_chaves ) :
    chaves = ''
    for x in lista_chaves :
        chaves += " row['%s'],"%x
    lista_ordenada  = sorted(lista_desordenada, key=lambda row:eval('(%s)'%(chaves)),reverse=False)
    return lista_ordenada