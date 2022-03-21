"""
----------------------------------------------------------------------------------------------
  BIBLIOTECA .: sql.py
  CRIACAO ....: 27/03/2021
  AUTOR ......: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO ..: Biblioteca de funcoes utilizadas para acessar o BD utilizado.

                - geraCnxBD()
                - geraCursorBD()
                - fechaCNxBD()
                - registraStatus()
              
----------------------------------------------------------------------------------------------
  HISTORICO ..: 
    * 27/03/2021 - Welber Pena de Sousa - Kyros Tecnologia
        - Criacao do script.
    
----------------------------------------------------------------------------------------------
"""

import cx_Oracle
import atexit
import datetime


class SQL() :
    def __init__(self, configuracoes) :
        try :
            log( '- Usuario .:', configuracoes.userBD if configuracoes.userBD else 'USUARIO DO BANCO DE DADOS NAO INFORMADO !!!' )
            log( '- Banco .:', configuracoes.banco )
            self.__cnx = cx_Oracle.connect( user = configuracoes.userBD, password = configuracoes.pwdBD, dsn = configuracoes.banco )
            self.__cursor = self.__cnx.cursor()
            self.__configuracoes = configuracoes
        except Exception as e :
            log('Erro ao conectar no banco : %s/%s@%s'%(configuracoes.userBD, '*'*(len(configuracoes.pwdBD)), configuracoes.banco))
            raise e


    def troca_owner(self, query):
        query = query.replace('gfcadastro.', self.__configuracoes.owner_gfcadastro.strip() + ( '.' if self.__configuracoes.owner_gfcadastro and not self.__configuracoes.owner_gfcadastro.strip().endswith('.') else '' ) )
        query = query.replace('GFCADASTRO.', self.__configuracoes.owner_gfcadastro.strip() + ( '.' if self.__configuracoes.owner_gfcadastro and not self.__configuracoes.owner_gfcadastro.strip().endswith('.') else '' ) )

        query = query.replace('gfcarga.', self.__configuracoes.owner_gfcarga.strip() + ( '.' if self.__configuracoes.owner_gfcarga and not self.__configuracoes.owner_gfcarga.strip().endswith('.') else '') )
        query = query.replace('GFCARGA.', self.__configuracoes.owner_gfcarga.strip() + ( '.' if self.__configuracoes.owner_gfcarga and not self.__configuracoes.owner_gfcarga.strip().endswith('.') else '') )
        
        query = query.replace('openrisow.', self.__configuracoes.owner_openrisow.strip() + ( '.' if self.__configuracoes.owner_openrisow and not self.__configuracoes.owner_openrisow.strip().endswith('.') else '') )
        query = query.replace('OPENRISOW.', self.__configuracoes.owner_openrisow.strip() + ( '.' if self.__configuracoes.owner_openrisow and not self.__configuracoes.owner_openrisow.strip().endswith('.') else '') )

        return query


    def description(self):
        if not self.__cnx :
            log("ERRO - Nao existe conexao ativa com banco de dados")
            raise "ERRO nao existe conexao com banco de dados"
        return  self.__cursor.description


    def executaRetorna(self, query):
        self.executa(query)
        return  self.fetchall()


    def executa(self, query, *args, **kwargs ):
        if not  self.__cnx :
            log("ERRO - Nao existe conexao ativa com banco de dados")
            raise "ERRO nao existe conexao com banco de dados"
        
        query = self.troca_owner(query)
        # print('ARGS ..', *args )
        if self.__configuracoes.ambiente == 'DEV' :
            log(query)
        try :
            self.__cursor.execute(query, *args, **kwargs)
        except Exception as e :
            log("Erro ao executar a query ... \n" + query)
            raise e
        
        return True


    def executaProcedure(self, procedure, *args, **kwargs ):
        """
        Funcao utilizada para executar uma procedure no banco :
            Os parametros podem ser passados de duas maneiras :
            1 - parametros sequenciais : 
                OWNER.PROCEDURE( 'param1', 'param2', 3 )
                   Ex: executaProcedure( 'OWNER.PROCEDURE', 'param1', 'param2', 3 )
            
            1 - parametros com nomes (chaves / key word ) :
                ** A ordem dos parametros nao importa pois sao identificados
                OWNER.PROCEDURE( p_2 => 'param1', p_1 = 'param2', indice = 3 )
                    Ex: executaProcedure( 'OWNER.PROCEDURE', indice = 3, p_2 = 'param1', p_1 = 'param2' )
        """
        if not  self.__cnx :
            log("ERRO - Nao existe conexao ativa com banco de dados")
            raise "ERRO nao existe conexao com banco de dados"

        proc = self.troca_owner(procedure)
        # print('ARGS ..', *args )
        log('Executando procedure', proc)
        if args :            
            log('- Com pararametros ..:\n   ', ',\n   '.join( str(x) for x in args ))
        elif kwargs :
            log('- Com pararametros ..:\n   ', ',\n   '.join( '%s = %s'%(k, kwargs[k]) for k in kwargs ))
            
        try:
            self.__cursor.callproc(proc, args, kwargs)
        except Exception as e :
            log('Erro ao executar a procedure %s'%(proc))
            raise e

        return True

    def var(self, tipo):
        if tipo == 'CLOB' :
            tipo = cx_Oracle.CLOB
        return self.__cursor.var(tipo)

    def fetchall(self) :
        return  self.__cursor.fetchall()
    
    
    def fetchone(self) :
        return  self.__cursor.fetchone()
    

    def fechaCnxDB(self) :
        log("Fechando conexao com o banco de dados .")
        if  self.__cursor :
             self.__cursor.close()
        if  self.__cnx :
             self.__cnx.close()
    

    def commit(self) :
        log('Realizando COMMIT !!!')
        self.__cnx.commit()

    def rollback(self) :
        log('Realizando rollback !!!')
        self.__cnx.rollback()
    
    def rowcount(self) :
        return self.__cursor.rowcount

def geraCnxBD(configuracoes) :
    log("Iniciando conexao com o banco de dados .")
    objeto = SQL(configuracoes)
    atexit.register(objeto.fechaCnxDB)
    return objeto
