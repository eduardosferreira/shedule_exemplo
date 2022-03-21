"""
----------------------------------------------------------------------------------------------
  BIBLIOTECA .: layout.py
  CRIACAO ....: 24/03/2021
  AUTOR ......: WELBER PENA DE SOUSA / KYROS TECNOLOGIA
  DESCRICAO ..: Biblioteca de funcoes utlizadas para tratar layouts de arquivos.

                    - carregaLayout()
                    - encodingDoArquivo()
                    - quebraRegistro()
              
----------------------------------------------------------------------------------------------
  HISTORICO ..: 
    * 24/03/2021 - Welber Pena de Sousa - Kyros Tecnologia
        - Criacao do script.
    
----------------------------------------------------------------------------------------------
"""

import os


def carregaLayout(layout = 'Todos', dir_layouts = None) :
    
    log('Carregando Layouts ....')
    if ( not dir_layouts ) or ( not os.path.isdir(dir_layouts) ) :
        dir_layouts = os.path.join('..', 'unificado', 'layout' )
        
    if not os.path.isdir(dir_layouts) :
        log('Erro - Diretorio de layouts nao encontrado !')
        log('- Diretorio layouts ..:', dir_layouts)
        return False

    if isinstance(layout, str) :
        if layout.lower() == 'todos' :
            layout = list(filter(lambda x: (x.lower().endswith('.csv') and x.lower().startswith('layout')) or x.lower().endswith('.layout'), os.listdir(dir_layouts) )) 
        else :
            if layout.__contains__(',') :
                layout = layout.split(',')
            else :
                layout = layout.split(' ')
    elif not isinstance(layout, (list, tuple)) :
        log('Erro - Tipo de layout citado nao permitido !')
        log('- Layout :', layout)
        return False
    
    dic_layouts = {}
    for arq_layout in layout :
        if not os.path.isfile(os.path.join( dir_layouts, arq_layout)) :
            log('ERRO - Arquivo de Layout nao existe ... < %s >'%arq_layout)
            log('Diretorio de layouts : ', os.path.realpath(dir_layouts))
            return False
        
        log('- Carregando layout : ', arq_layout)
        fd = open(os.path.join(dir_layouts, arq_layout), 'r')
        y = 0
        nome_layout = arq_layout.split('.')[0]
        dic_layouts[nome_layout] = {}
        dic_layouts[nome_layout]['dic_campos'] = {}
        dic_registros = dic_layouts[nome_layout]
        dic_campos = dic_registros['dic_campos']
        tipo_layout = False
        separador = False
        separador_de_campos = False
        
        for item in fd.readlines() :
            if item.strip().replace('\n','') and not item.startswith('#') :
                if not separador :
                    separador = ';' if item.__contains__(';') else ','
                
                y += 1
                if not tipo_layout or tipo_layout == 'posicional' :
                    linha = item.replace('\n','').split(separador)
                    #### Formato do layout :
                    ### Ordem do campo;NomeCampo;Tamanho,Inicio,Fim
                    ### 1;CD_EMPRESA;5;1;5
                elif tipo_layout == 'csv' :
                    linha = item.replace('\n','').split(separador_de_campos)

                if not linha[0].isdigit() :
                    linha = [str(y)] + linha

                dic_registros[int(linha[0])] = []

                if len(linha) < 5 and not tipo_layout :
                    linha_temp = linha[:2]
                    linha_temp.append(0)
                    linha_temp += linha[2:]
                    linha = linha_temp
                elif len(linha) < 5 and tipo_layout in [ 'csv', 'posicional' ] :
                    while len(linha) < 5 :
                        linha.append(0)
                    # if y == 1 :
                    #     linha[3] = '1'
                    # else :
                    #     linha[3] = str(int(dic_registros[y-1][3]) + 1)
                    # linha[4] = str(int(linha[3]) + int(linha[2]) -1)
                if not tipo_layout and not linha[2] :
                    linha[2] = str((int(linha[4]) - int(linha[3])) + 1 )

                for r in linha[1:] :
                    if isinstance(r, str ) and r.isdigit() :
                        dic_registros[int(linha[0])].append(int(r))
                    else :
                        dic_registros[int(linha[0])].append(r)

                dic_campos[linha[1]] = int(linha[0])
                # print(';'.join(  str(x) for x in dic_registros[int(linha[0])]  ) )
                # log('> ',y, '=',dic_registros[int(linha[0])])
                # log('> ',linha[0], '=',dic_registros[int(linha[0])])
                # log(linha[1],'=',linha[0])
            else :
                if item.__contains__('=') :
                    if item.split('=')[0].__contains__('TipoLayout') :
                        tipo_layout = item.split('=')[-1].strip().lower()
                        log('- Tipo de Layout :', tipo_layout)
                        dic_layouts[nome_layout]['tipo_layout'] = tipo_layout
                        if tipo_layout == 'posicional' :
                            separador = ';'
                    elif item.split('=')[0].__contains__('SeparadorDeCampos') :
                        separador_de_campos = item.split('=')[-1].strip()
                        log('- Separador de campos :', separador_de_campos)
                        dic_layouts[nome_layout]['separador_de_campos'] = separador_de_campos

    log('Dicionario de Layouts criado !')
    
    carregaLayout.dic_layouts = dic_layouts
    return dic_layouts


def encodingDoArquivo(path_arq) :
    try :
        fd = open(path_arq, 'r', encoding='utf-8')
        fd.read()
        fd.close()
    except :
        return 'iso-8859-1'

    return 'utf-8'


def quebraRegistro(reg, dic_registros) :
    # print(carregaLayout.dic_layouts['FAT'].keys())
    # print(carregaLayout.dic_layouts.keys())
    if isinstance( dic_registros, str ) and dic_registros in carregaLayout.dic_layouts.keys() :
        dic_registros = carregaLayout.dic_layouts[dic_registros]
    
    tipo_layout = dic_registros.get('tipo_layout', False)
    separador_de_campos = dic_registros.get('separador_de_campos', False)
    # dic_campos = dic_registros['dic_campos']
    itens_registro = []
    colunas = []
    if not tipo_layout :
        for y in range( 1, len(dic_registros.keys())+1 ) :
            field, t, i, f = dic_registros[y]
            itens_registro.append( reg[i-1:f] )
            colunas.append(field)
            print(field, '=', reg[i-1:f])
    
    else :
        # print( len(dic_registros['dic_campos'].keys()), len(dic_registros.keys()) )
        if tipo_layout == 'csv' :
            for y in range( 1, len(dic_registros['dic_campos'].keys())+1 ) :
                # print( y)
                if len(reg.split(separador_de_campos)) >= y :
                    field = dic_registros[y][0]
                    itens_registro.append( reg.split(separador_de_campos)[y-1] )
                    colunas.append(field)
                    print(y, field, '=', itens_registro[y-1])
        
        elif tipo_layout == 'posicional' :
            for y in range( 1, len(dic_registros['dic_campos'].keys())+1 ) :
                # print( y)
                field, t, i, f = dic_registros[y]
                itens_registro.append( reg[i-1:f] )
                
                # field = dic_registros[y][0]
                # itens_registro.append( reg.split(separador_de_campos)[y-1] )
                colunas.append(field)
                # print(y, field, '=', itens_registro[y-1])

    # print(colunas)

    return itens_registro


def geraLinha( dados, layout ) :
    linha = ""
    dic_registros = carregaLayout.dic_layouts[layout]
    if dic_registros['tipo_layout'].lower() == 'posicional' :
        ### Busca o maior campo do layout .
        qtd_campos = 0
        for k in dic_registros.keys() :
            if isinstance(k, (int,float) ) or k.isdigit() :
                if k > qtd_campos :
                    qtd_campos = k
        qtd_campos += 1

        for k in range(1,qtd_campos) :
            campo = dic_registros[k][0]
            tam = dic_registros[k][1]
            # print(campo, tam, dados[campo])
            if dados[campo].isdigit() :
                linha += dados[campo].rjust( tam,' ' )[-tam:]
            else :
                linha += dados[campo][:tam].ljust(tam, ' ')

    return linha
