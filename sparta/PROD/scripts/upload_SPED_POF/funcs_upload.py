###### !/usr/local/bin/python3.7
###### -*- coding: utf-8 -*-

"""
----------------------------------------------------------------------------------------------
  SISTEMA ..: 
  MODULO ...: 
  SCRIPT ...: funcs_upload.py
  CRIACAO ..: 21/05/2020
  AUTOR ....: WELBER PENA DE SOUSA / KYROS CONSULTORIA
  DESCRICAO : 
                

  ANEXO ....: Demais dados e documentação na pasta documentação, arquivos :
                - Teshuva_EspecificaçãoFuncionalProcUpload_EliminarAçãoManual_v1.docx
                - Teshuva_EspecificaçãoFuncionalProcUpload_EliminarAçãoManual_v2.docx

----------------------------------------------------------------------------------------------
  HISTORICO : 
    * 21/05/2020 - Welber Pena de Sousa - Kyros Consultoria
        - Criacao do script.
    * 09/09/2021 Victor Santos - Kyros Consultoria
      Adequação ao novo formato de Python
----------------------------------------------------------------------------------------------
"""

import os
import sys

global SD, dir_base
SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV'+SD)[0], 'DEV')
sys.path.append(dir_base)
import configuracoes

import fnmatch

import time
import base64
from urllib.parse import urlencode, quote_plus
import http
import requests
import urllib.parse
import uuid
import hashlib
import json
import comum

comum.carregaConfiguracoes(configuracoes)

def execUpload( dic_args = None, autenticacao = None, data = None, post = True ) :
    funcao = 'upload'
    if not autenticacao :
        autenticacao = authenticate()
    
    if autenticacao :
        resultado = {'retorno': False}
        port = int(configuracoes.porta_webservice)
        host = configuracoes.ip_webservice 
        path_pagina = 'http://%s:%s/pof/api/%s'%( host, port, funcao )
               
        
        #conn = http.HTTPConnection( host )
        # conn = http.client.HTTPConnection( host, port )
                
        # log( "Executando %s"%(funcao))
        if type(dic_args) == dict and len(dic_args.keys()) > 0 :
            #### getDadosConexaoDB.json?id_auth_user=%s&id_tns=%s&id_user_db=%s
            # log("--- DIC_ARGS --->", dic_args['id_atualizador'])
            path_pagina += '?'
            path_pagina += urllib.parse.urlencode(dic_args)
        
        headers = {
            # 'Content-Type': 'application/json' 
            # 'Content-Type': 'application/octet-stream' 
            'Content-Type': 'text/x-log' 
        }
        # if data :
        #     headers['Content-Length'] = str(len(data))
        if autenticacao :
            # log('><><>', autenticacao.headers.keys())
            headers['x-xsrf-token'] = autenticacao.get('XSRF-TOKEN', False)
        
        if autenticacao.get('Cookie', False) :
            headers['Cookie'] = autenticacao['Cookie']
        
#        log('url execUpload => :', path_pagina)
        # log('\nheaders :', headers)
        # log('\ndata :', data)

        # try :
        if True :
            if post :
                response = requests.post(path_pagina, headers=headers, data=data )
            else :
                response = requests.get(path_pagina, headers=headers, data=data )
        # except ConnectionRefusedError :
        #     log("Erro ao tentar conectar ao ip : %s"%(host))
        #     return False

        # r1 = conn.getresponse()
#        log( 'text ...... =', response.text )
#        log( 'reason .... =', response.reason )
#        log( 'json ...... =', response.json )
#        log( 'status_code =', response.status_code )
        
        if response and  response.status_code == 200:
            # log(dir(response))
            
            response_json = response.json
            # log ("########## DEBUG ##########")
            # log ("########## DEBUG ##########")
            # log ("########## DEBUG ##########")
            # log ("Resposta do execUpload = ",response_json )
            # log ("########## DEBUG ##########")
            # log ("########## DEBUG ##########")
            
            if response.text == 'All finished.' :
                return 'Finished'

            return True
        
        if response.status_code == 404 and not post :
            return True

        log( 'ERRO %s \n Content : %s'%(response.status_code, response.content) )
    
    return False


def authenticate():
    log('Autenticando ... ')
    username = configuracoes.usuario_webservice
    secret = configuracoes.senha_webservice

    port = int(configuracoes.porta_webservice)

    data = {
        'username': username,
        'password': secret,
    }

    headers = {
        'content-type' : 'application/json'
    }
    
    autenticou = False
    response = False

    ip_webservice = configuracoes.ip_webservice
    # ip_webservice = config['ip_webservice'] 
    auth_url = 'http://%s:%s/pof/api/login'%( ip_webservice, port )
    try :
        response = requests.post(auth_url, headers=headers, json=data)
    except :
        log( 'ERRO - Authentication failed: Erro na chamada do webservice %s !'%(ip_webservice) )

    if response and  response.status_code == 200:
        response_json = response.json()
        # log(response_json.get('access_token'))
        if 'displayName' in response_json :
            autenticou = response

    if not autenticou :
        log( 'ERRO - Authentication failed: %s \n Content : %s'%(response.status_code, response.content) )
        return False
    log('- Autenticacao - OK - %s'%(autenticou.json().get('displayName')))

    autenticou = autenticou.json()
    for super_item in response.headers['Set-Cookie'].split(';') :
        if super_item :
            for item in super_item.split(','):
                if item.__contains__('=') :
                    i, valor = item.split('=')
                    if i == 'XSRF-TOKEN' :
                        autenticou['XSRF-TOKEN'] = valor
    
    cookie = ''
    for item in response.cookies.items():
        if cookie :
            cookie += '; '
        cookie += '%s=%s'%(item)
    autenticou['Cookie'] = cookie

    return autenticou


def arquivoAEnviar(arq, mascaras) :
    valido = False
    for mascara in mascaras :
        if fnmatch.fnmatch(arq,mascara) :
            valido = True
    return valido


def enviarArquivo(autenticacao, dir_trabalho, arq, lst_partes = None, guuid = None, tentativas = 0 ) :
    x = 1
    tam_arquivo = os.stat( os.path.join(dir_trabalho, arq)).st_size

########## alterado para binario - airton 20201202    fd = open(os.path.join(dir_trabalho, arq), 'r', encoding=encodingDoArquivo(os.path.join(dir_trabalho, arq)))
    fd = open(os.path.join(dir_trabalho, arq), 'rb')
   
    md5 = hashlib.md5()
    ### chunkSize: 4*1024*1024,
    tam_txt = 4*1024*1024
    # tam_arquivo = 2*tam_txt

 #   partes = round(tam_arquivo / tam_txt) if tam_arquivo / tam_txt <= round(tam_arquivo / tam_txt) else round(tam_arquivo / tam_txt) +1

    partes = int(tam_arquivo / tam_txt)
    if( partes == 0 ):
        partes = 1

        
    log('- Tamanho do arquivo ..: %s bytes'%( tam_arquivo ))
    log('- Tamanho dos envios ..: %s bytes'%( tam_txt ))
    log('- Qtde. de envios .....: %s'%( partes ))
    
    if not guuid :
        identificador_uuid = str(uuid.uuid4())
    else :
        identificador_uuid = guuid
    log('- ID unico ............: %s'%( identificador_uuid ))
    log('Aguardando envio ...')
    
    t = 0
    dic_args = {}
    
    if not lst_partes :
        lst_partes = range(1,partes+1)

    finalizado_OK = False
  
    for x in range(1, partes+1) :

        if (x < partes):
            txt = fd.read(tam_txt)
        else: 
            txt = fd.read(2*tam_txt)
    
        tamtxt = len(txt)
            
########### alterado para binario - airton 20201202            md5.update(txt.encode())
        tam = 0
        t += len(txt)
        
        if x in lst_partes :
            # log( 'Enviando:', x, t)
            dic_args['resumableChunkSize'] = str(tam_txt)
            dic_args['resumableTotalSize'] = str(tam_arquivo)
            
            dic_args['resumableChunkNumber'] = x
            dic_args['resumableTotalChunks'] = partes
            dic_args['resumableIdentifier'] = identificador_uuid
            
            tam = str(len(txt))

            dic_args['resumableCurrentChunkSize'] = tamtxt
            dic_args['resumableCurrentChunkSize'] = tam
  
            dic_args['resumableFilename'] = arq
            dic_args['resumableRelativePath'] = arq
#            log("  = ", )
#            log("  = ", )
#            log("dic_args  = ", dic_args )
#            log("  = ", )
#            log("  = ", )
            ret = execUpload( dic_args=dic_args, autenticacao=autenticacao, data=txt ) 
#            log("retorno = ", ret)
#            log("  = ", )
#            log("  = ", )
          
            if not ret :
                log('ERRO - ao realizar upload do arquivo.')
                return False
            finalizado_OK = True if ret == 'Finished' else False
        # log( res)
    dic_args['hash-md5'] = md5.hexdigest()
    fd.close()
    log('- Upload OK')

    erros = []
    x = 1
    if not finalizado_OK :
        log('Validando partes enviadas ...')
        for x in range(1,partes+1) :
            if x in lst_partes :
                # log('Validando', x)

                dic_args['resumableChunkSize'] = str(tam_txt)
                dic_args['resumableTotalSize'] = str(tam_arquivo)
                
                dic_args['resumableChunkNumber'] = x
                dic_args['resumableTotalChunks'] = partes
                dic_args['resumableIdentifier'] = identificador_uuid
                
                dic_args['resumableCurrentChunkSize'] = tam_txt

                if (x == partes):
                    dic_args['resumableCurrentChunkSize'] = tam
                               
                dic_args['resumableFilename'] = arq
                dic_args['resumableRelativePath'] = arq
                if not execUpload( dic_args=dic_args, autenticacao=autenticacao, post=False ) :
                    log('- Parte %s com problema, marcando para reenviar.'%(x))
                    erros.append(x)
    if erros :
        if tentativas < 10 :
            return enviarArquivo( autenticacao, dir_trabalho, arq, erros, identificador_uuid, tentativas + 1 )
        else :
            log('ERRO - Numero de tentativas < %s > excedidas para o arquivo.'%(tentativas))
            return False
    else :
        log('- Arquivo enviado com sucesso ... %s partes'%(x))

    return dic_args


def registraUpload( autenticacao, id_obrigacao, lst_protocolados = [], lst_recibos = [], lst_regerados = [], lst_analise = [] ) :
    log('Registrar upload do arquivo.')
    url = '/pof/api/teshuva/controle/'
    
    idUnidadeOrganizacional = identificaUnidadeOrganizacional( autenticacao, getattr(configuracoes, 'empresa', False), getattr(configuracoes, 'filial', False)) 
    # idUnidadeOrganizacional = 2
    
    if not idUnidadeOrganizacional :
        log('ERRO - Não encontrado o idUnidadeOrganizacional .')
        return False
    
    if id_obrigacao == 1 :
        extensao = configuracoes.serie_original
    else :
        extensao = ''
    
    periodoId = "%s%s"%( configuracoes.ano, configuracoes.mes )

    id_controle = identificaIdControle( autenticacao, id_obrigacao, idUnidadeOrganizacional, extensao, periodoId )


    log("id_controle = ",id_controle )

    if id_controle == False :
        id_controle = "370"

    data = {}
    data["idObrigacao"] = id_obrigacao     ### {id_obrigacao}, // a lista de possíveis valores pode ser obtida com GET em /api/obrigacoes)
    data["idUnidadeOrganizacional"] = idUnidadeOrganizacional   ### {id_organizacao}, // a lista de possíveis valores pode ser obtida com GET em /api/unidadesOrganizacionais

    if id_obrigacao == 1 :
        data["extensao"] = configuracoes.serie
    else :
        data["extensao"] = ""                  ### "{serie_ou_uf}", // obrigações com o Convenio 115 e Convênio 201, informar a série. Para obrigações sem quebra por série ou UF (Como GIA e SPED) informar "" (vazio)
    
    data["originais"] = lst_protocolados     ### [{FileRepresentation}], // Veja a definição de um objeto FileRepresentation mais abaixo. Informar a lista de arquivos Orginais (protocolados). Se não existir, informar como um array vazio.
    data["regerados"] = lst_regerados        ### [{FileRepresentation}], // Veja a definição de um objeto FileRepresentation mais abaixo. Informar a lista de arquivos de candidatos do Teshuvá (regerados). Se não existir, informar como um array vazio.
    data["analises"] = lst_analise           ### [{FileRepresentation}], // Veja a definição de um objeto FileRepresentation mais abaixo. Informar a lista de arquivos de Análise (análise teshuvá). Se não existir, informar como um array vazio.
    data["recibos"] = lst_recibos            ### [{FileRepresentation}], // Veja a definição de um objeto FileRepresentation mais abaixo. Informar a lista de arquivos de Recibo (recibos dos protocolados). Se não existir, informar como um array vazio.
   
    data["originaisPreSelected"] = "true"     ### false, // valor fixo
    data["regeradosPreSelected"] = "true"   ### false, // valor fixo
    data["analisesPreSelected"] = "true"    ### false, // valor fixo,
    data["recibosPreSelected"] = "true"     ### false, // valor fixo

    if (lst_protocolados):
        data["originaisPreSelected"] = "false"
    if (lst_regerados):
        data["regeradosPreSelected"] = "false"
    if (lst_analise):
        data["analisesPreSelected"] = "false"        
    if (lst_recibos):
        data["recibos"] = "false"        
    
    data["periodo"] = periodoId
    
    data["usuarioLoginAD"] = autenticacao['displayName']
    
    data["entregaSemOriginais"] = "false"
    if lst_recibos :
        data["entregaSemOriginais"] = "true"         ### {true|false} // Informar True se não há arquivos protocolados e/ou recibos, caso contrário informar false. Se for informado false e os arrays originais ou recibos estiverem vazios, um erro será gera   
    
    if id_controle :
        url += '%s'%(id_controle)
    
    if not autenticacao :
        autenticacao = authenticate()
    
    if autenticacao :
        resultado = {'retorno': False}
        port = int(configuracoes.porta_webservice)
        host = configuracoes.ip_webservice 
        path_pagina = 'http://%s:%s%s'%( host, port, url )
        
        headers = {
            'Content-Type': 'application/json' 
        }

        headers['X-XSRF-TOKEN'] = autenticacao.get('XSRF-TOKEN', False)
        
        if autenticacao.get('Cookie', False) :
            headers['Cookie'] = autenticacao['Cookie']
        try :
            if id_controle :
#                log('=====>>>> P U T <<<<=====')
                response = requests.put(path_pagina, headers=headers, json=data )
            else :
#                log('=====>>>> P O S T <<<<=====')
                response = requests.post(path_pagina, headers=headers, json=data )
        except ConnectionRefusedError :
            log("Erro ao tentar conectar ao ip : %s"%(host))
            return False
        
        if response and response.status_code == 200:
            response_json = response.json
            return True

        log( 'ERRO %s \n Content : %s'%(response.status_code, response.content) )

    return False


def identificaUnidadeOrganizacional( autenticacao, empresa, filial ) :
    url = "/pof/api/unidadesOrganizacionais/codigos"

    headers = {}
    headers['content-type'] = 'application/json'
    headers['x-xsrf-token'] = autenticacao.get('XSRF-TOKEN', False)
    if autenticacao.get('Cookie', False) :
        headers['Cookie'] = autenticacao['Cookie']
        
    resultado = {'retorno': False}
    port = int(configuracoes.porta_webservice)
    host = configuracoes.ip_webservice 
    path_pagina = 'http://%s:%s%s'%( host, port, url )
    try :
        response = requests.get(path_pagina, headers=headers )
    except ConnectionRefusedError :
        log("Erro ao tentar conectar ao ip : %s"%(host))
        return False

    # r1 = conn.getresponse()

    if response and  response.status_code == 200:
        ret = False
        for item in response.json() :
            emp = item.get('organizacaoCodigo', '')
            if emp.__contains__(empresa) and emp.__contains__(filial) :
                ret = item['organizacaoId']

        return ret

    log( 'ERRO %s \n Content : %s'%(response.status_code, response.content) )

    return False


def identificaIdControle( autenticacao, id_obrigacao, idUnidadeOrganizacional, extensao, periodoId ) :

    url = "/pof/api/teshuva/controle/%s/processos"%(id_obrigacao)
    log('Identifica idControle ...')
    
    log(" = ", )
    log(" = ", )
    log(" = ", ) 
    log("autenticacao = ",autenticacao )
    log("id_obrigacao = ",id_obrigacao )
    log("idUnidadeOrganizacional = ",idUnidadeOrganizacional )
    log("extensao = ",extensao )
    log("periodoId = ",periodoId )
    log("url = ",url )
    log(" = ", )
    log(" = ", )
    log(" = ", )
    log(" = ", )

    headers = {}
    headers['content-type'] = 'application/json'
    headers['x-xsrf-token'] = autenticacao.get('XSRF-TOKEN', False)
    if autenticacao.get('Cookie', False) :
        headers['Cookie'] = autenticacao['Cookie']
        
    resultado = {'retorno': False}
    port = int(configuracoes.porta_webservice)
    host = configuracoes.ip_webservice 
    path_pagina = 'http://%s:%s%s'%( host, port, url )

    dic_parametros = {}
    dic_parametros['periodoID'] = periodoId
    dic_parametros['organizacaoId'] = idUnidadeOrganizacional
    dic_parametros['extensao'] = extensao

    path_pagina += '?'
    path_pagina += urllib.parse.urlencode(dic_parametros)

#    log('url :', path_pagina)
    # log('\nheaders :', headers)
    # log('\ndata :', data)

    try :
        response = requests.get(path_pagina, headers=headers )
    except ConnectionRefusedError :
        log("Erro ao tentar conectar ao ip : %s"%(host))
        return False

    # r1 = conn.getresponse()
    ret = False
    if response and  response.status_code == 200:
        response_json = response.json()
#        log('XXXX', response_json, 'XXXXXX')
        # log(dir(response))
        
        if response_json['numberOfElements'] > 0 :
            item = response_json['content'][0]
            # log(item)
            ret = item.get('teshuvaControleId', '')
    else :
        log( 'ERRO %s \n Content : %s'%(response.status_code, response.content) )

    log('- IdControle = %s'%(ret))
    return ret


def executaSQL( autenticacao, query, size = 0, page = 0 ) :
    url = "/pof/generic/query"
    log('Executa Query sql ...')

    headers = {}
    headers['content-type'] = 'application/json'
    headers['x-xsrf-token'] = autenticacao.get('XSRF-TOKEN', False)
    if autenticacao.get('Cookie', False) :
        headers['Cookie'] = autenticacao['Cookie']
        
    resultado = {}
    port = int(configuracoes.porta_webservice)
    host = configuracoes.ip_webservice 
    path_pagina = 'http://%s:%s%s'%( host, port, url )

    data = {}
    data['page'] = page
    data['size'] = size
    data['sql'] = query

#    log('url :', path_pagina)
    # log('\nheaders :', headers)
    # log('\ndata :', data)

    try :
        response = requests.post(path_pagina, headers=headers, data = json.dumps( data) )
    except ConnectionRefusedError :
        log("Erro ao tentar conectar ao ip : %s"%(host))
        return False

    # r1 = conn.getresponse()
    ret = False
    if response and  response.status_code == 200:
        response_json = response.json()
#        log('###### RESPONSE_JASON = ', response_json, '######')
        # log(dir(response))
        resultado = response_json
        # log( 'text =', response.text )
        # log( 'reason =', response.reason )
#        log( 'json =', response.json )
        # response_json = response.json
    else :
        log( '###### ERRO %s \n Content : %s'%(response.status_code, response.content) )

    return resultado



