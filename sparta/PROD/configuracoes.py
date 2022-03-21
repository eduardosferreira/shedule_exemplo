#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-

import socket
import sys
import os

SD = ('/' if os.name == 'posix' else '\\')
dir_base = os.path.join( os.path.realpath('.').split(SD+'PROD'+SD)[0], 'PROD') if os.path.realpath('.').__contains__(SD+'PROD'+SD) else os.path.join( os.path.realpath('.').split(SD+'DEV')[0], 'DEV')

sys.path.append( os.path.join( dir_base, 'lib'))

# dir_base = os.path.join( os.path.realpath('.').split('/PROD/')[0], 'PROD') if os.path.realpath('.').__contains__('/PROD/') else os.path.join( os.path.realpath('.').split('/DEV/')[0], 'DEV')
ambiente = dir_base.split('/')[-1]

raiz = '' if ambiente == 'PROD' else os.path.realpath('.')


def retornaIpLocal():
    st = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:       
        st.connect(('10.255.255.255', 1))
        IP = st.getsockname()[0]
    except Exception:
        IP = '127.0.0.1'
    finally:
        st.close()
    return IP


ip = retornaIpLocal()

if not ip :
    print(ip)
    raise 'IP da maquina invalido ...'


# print('DIR_BASE >>>', dir_base)
# Banco para a Clone 1 = GFCLONEDEV
# Banco para a Clone 2 = GFCLONEPREPROD
# Banco para a Clone 6 = GFPRODC6
# Banco para a Clone 7 = GFPRODC7
##### Configuracoes :
maquina = 'local'
banco = "GFCLONEDEV"

if ip == '10.238.10.208' :
    maquina = 'clone1'
    banco = 'GFCLONEDEV'
    
elif ip == '10.238.10.209' :
    maquina = 'clone2'
    banco = "GFCLONEPREPROD"

elif ip == '10.238.10.210' :
    maquina = 'clone6'
    banco = 'GFPRODC6'

elif ip == '10.238.10.109' :
    maquina = 'clone7'
    banco = 'GFPRODC7'

raiz = '' if ambiente == 'PROD' else os.path.join( dir_base, 'arquivos', maquina, os.path.realpath('.').split('/scripts/')[-1].split('/')[-1] )

dir_log = [ dir_base, 'log', maquina ] + os.path.realpath('.').split('/scripts/')[-1].split('/')
dir_log = os.path.join( *dir_log )

dir_geracao_arquivos = [ dir_base, 'saidas', maquina ] + os.path.realpath('.').split('/scripts/')[-1].split('/')
dir_geracao_arquivos = os.path.join( *dir_geracao_arquivos )
dir_entrada = os.path.join( dir_base, 'entradas', os.path.realpath('.').split('/')[-1] )

if ambiente == 'PROD' :
    owner_gfcadastro = 'gfcadastro.'
    owner_gfcarga = 'gfcarga.'
    owner_openrisow = 'openrisow.'
    dir_entrada_carga = '/portaloptrib/TESHUVA/POC_Cadastro_Billing'
    dir_arquivos_protocolados  = '/portaloptrib/ARQUIVOS_PROTOCOLADO'
    fd = open(os.path.join(dir_base, '.userBase.ini'), 'r')
    userBD, pwdBD = fd.readline().replace('\r','').replace('\n','').split('/')


else :
    # owner_gfcadastro = 'gfcadastro.'
    # owner_gfcarga = 'gfcarga.'
    # owner_openrisow = 'openrisow.'
    owner_gfcadastro = ''
    owner_gfcarga = ''
    owner_openrisow = ''
    dir_entrada_carga = os.path.join( dir_base, 'entradas', os.path.realpath('.').split('/')[-1] )
    dir_arquivos_protocolados  = os.path.join( dir_base, 'entradas', os.path.realpath('.').split('/')[-1],'ARQUIVOS_PROTOCOLADO' )
    userBD = ''
    pwdBD = ''














