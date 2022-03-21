import fcntl
import array
import struct
import socket
import platform

import sys
from functools import wraps
import platform
import xmlrpc.client
import json
import http

def retornaIpsLocal():
    SIOCGIFCONF = 0x8912
    MAXBYTES = 8096
    arch = platform.architecture()[0]
    var1 = -1
    var2 = -1
    if arch == '32bit':
        var1 = 32
        var2 = 32
    elif arch == '64bit':
        var1 = 16
        var2 = 40
    else:
        raise OSError("Unknown architecture: %s" % arch)
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    names = array.array('u', b' ' * MAXBYTES)
    outbytes = struct.unpack('iL', fcntl.ioctl(
        sock.fileno(),
        SIOCGIFCONF,
        struct.pack('iL', MAXBYTES, names.buffer_info()[0])
        ))[0]
    namestr = names.tobytes()
    dic_ips = {}
    for i in range(0, outbytes, var2) :
        dic_ips[namestr[i:i+var1].split(b'\0', 1)[0].decode()] = socket.inet_ntoa(namestr[i+20:i+24])
    return dic_ips


class TimeoutTransport(xmlrpc.client.Transport):
    timeout = 20.0
    def set_timeout(self, timeout):
        self.timeout = timeout
    def make_connection(self, host):
        h = http.client.HTTPConnection(host, timeout=self.timeout)
        return h


def execFncXmlRpc(cfg, funcao, *param ):
    if param :
        comando = "json.loads(cnx_gtw.{}{})".format(funcao, param )
    else :
        comando = "json.loads(cnx_gtw.{}())".format(funcao)
    ret = ['ERRO']
    try:
        transporte = TimeoutTransport()
        transporte.set_timeout(20.0)
        str_server = "http://{}:{}/".format(
            cfg["HOST"], cfg["PORT"])
        cnx_gtw = xmlrpc.client.ServerProxy(
            str_server, allow_none=True, transport=transporte)  # , verbose=True)
        # print('COMANDO >>>>>', comando)
        ret = eval(comando)
        loop = False
    except xmlrpc.client.Fault as err:
        log("""Falha ocorreu
                Metodo.......: {}
                Parametros...: {}
                Fault code...: {}
                Fault string.: {}""".format(funcao, param, err.faultCode, err.faultString))
        ret.append(err.faultString)
    except xmlrpc.client.ProtocolError as err:
        log("""Erro de protocolo
                URL.................: {}
                HTTP/HTTPS headers..: %s
                Error code..........: %d
                Error message.......: %s""".format(err.url, err.headers, err.errcode, err.errmsg))
        ret.append(err.errmsg)
    except:
        log("Erro generico")
        log(('execFncXmlRpc {} - {}'.format(funcao, sys.exc_info())))
        ret.append('Erro generico')
    cnx_gtw._ServerProxy__close()
    return ret
