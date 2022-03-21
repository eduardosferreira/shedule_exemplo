# -*- coding: utf-8 -*-
"""
Created on Tue Jan 26 10:26:00 2021

@author: Airton
"""

def sonum(parametro):
    retorno = "" 
    for i in range(len(parametro)):
        if (parametro[i].isdigit()):
            retorno = retorno + parametro[i]
    return(retorno)

def valida_insc_est(estado,insc):
    inscsp = insc
   
    if( estado.upper() in ('SP','RJ','ES','MG','PR','SC','RS','DF','GO','TO','MT','MS','AC','RO','BA','SE','PE','AL','PB','RN','CE','PI','PA','AM','RR','AP','MA','TO')):
        estado = estado.upper()
    else:
        return False

    if ( ( insc.upper() == 'ISENTO') or (insc.strip()== '') ): 
        return('0')
    else:
        insc=sonum(insc)
        if len(insc) == 0:
            return False
        if int(insc) == 0:
            return False

################################################
#
################################################

    if estado == 'AC':
        if len(insc) != 13:
            return False
        if insc[0:2] != '01':
            return False

        v_nr_calculo = int(insc[0]) * 4
        v_nr_calculo = v_nr_calculo +  int(insc[1]) * 3
        v_nr_calculo = v_nr_calculo +  int(insc[2]) * 2
        v_nr_calculo = v_nr_calculo +  int(insc[3]) * 9
        v_nr_calculo = v_nr_calculo +  int(insc[4]) * 8
        v_nr_calculo = v_nr_calculo +  int(insc[5]) * 7
        v_nr_calculo = v_nr_calculo +  int(insc[6]) * 6
        v_nr_calculo = v_nr_calculo +  int(insc[7]) * 5
        v_nr_calculo = v_nr_calculo +  int(insc[8]) * 4
        v_nr_calculo = v_nr_calculo +  int(insc[9]) * 3
        v_nr_calculo = v_nr_calculo +  int(insc[10]) * 2
        v_nr_digito = ( 11 - v_nr_calculo%11) 
        if int(v_nr_digito) >= 10:
            v_nr_digito = 0
           
        if int(v_nr_digito) != int(insc[11]):
            return False
        else:
            v_nr_calculo = int(insc[0]) * 5
            v_nr_calculo = v_nr_calculo +  int(insc[1]) * 4
            v_nr_calculo = v_nr_calculo +  int(insc[2]) * 3
            v_nr_calculo = v_nr_calculo +  int(insc[3]) * 2
            v_nr_calculo = v_nr_calculo +  int(insc[4]) * 9
            v_nr_calculo = v_nr_calculo +  int(insc[5]) * 8
            v_nr_calculo = v_nr_calculo +  int(insc[6]) * 7
            v_nr_calculo = v_nr_calculo +  int(insc[7]) * 6
            v_nr_calculo = v_nr_calculo +  int(insc[8]) * 5
            v_nr_calculo = v_nr_calculo +  int(insc[9]) * 4
            v_nr_calculo = v_nr_calculo +  int(insc[10]) * 3
            v_nr_calculo = v_nr_calculo +  int(insc[11]) * 2
            v_nr_digito = ( 11 - v_nr_calculo%11) 
            if int(v_nr_digito) >= 10:
                v_nr_digito = 0
           
            if int(v_nr_digito) != int(insc[12]):
                return False

################################################
#
################################################

    elif estado == 'AL':
        if len(insc) != 9:
            return False
        if insc[0:2] != '24':
            return False
        # if ( not insc[2] in ('0','3','5','7','8')):
        #     return False

        v_nr_calculo = int(insc[0]) * 9
        v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
        v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
        v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
        v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
        v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
        v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
        v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
        v_nr_digito = ( v_nr_calculo * 10 ) - ( int((v_nr_calculo * 10) / 11) * 11 )

        if int(v_nr_digito) == 10 :
            v_nr_digito = 0
        
        if int(v_nr_digito) != int(insc[8]):
            return False

################################################
#
################################################

    elif estado == 'AP':
        
        if len(insc) != 9:
            return False
        if insc[0:2] != '03':
            return False

        if insc >= '03000001' and  insc <= '03017000':
            v_nr_calculo = 5
            v_nr_digito_aux_1 = 0
            
        if insc >= '03017001' and  insc <= '03019022':
            v_nr_calculo = 9
            v_nr_digito_aux_1 = 1    
            
        if insc >= '03019023':
            v_nr_calculo = 0
            v_nr_digito_aux_1 = 0    
      

        v_nr_calculo = v_nr_calculo + (int(insc[0]) * 9)
        v_nr_calculo = v_nr_calculo + (int(insc[1]) * 8)
        v_nr_calculo = v_nr_calculo + (int(insc[2]) * 7)
        v_nr_calculo = v_nr_calculo + (int(insc[3]) * 6)
        v_nr_calculo = v_nr_calculo + (int(insc[4]) * 5)
        v_nr_calculo = v_nr_calculo + (int(insc[5]) * 4)
        v_nr_calculo = v_nr_calculo + (int(insc[6]) * 3)
        v_nr_calculo = v_nr_calculo + (int(insc[7]) * 2)
        v_nr_digito = ( 11 - (v_nr_calculo%11) )
        
        if int(v_nr_digito) == 10:
            v_nr_digito = 0

        if int(v_nr_digito) == 11:
            v_nr_digito = v_nr_digito_aux_1
       
        if int(v_nr_digito) != int(insc[8]): 
            return False
 
################################################
#
################################################

    elif estado == 'MG':
        if len(insc) != 13:
            return False
        else:
        
            v_st_calculo = insc[0:3] + '0' +  insc[3:]
            
            v_calculo = str((int(v_st_calculo[0]) * 1))
            v_calculo = v_calculo + str((int(v_st_calculo[1]) * 2))
            v_calculo = v_calculo + str((int(v_st_calculo[2]) * 1))
            v_calculo = v_calculo + str((int(v_st_calculo[3]) * 2))
            v_calculo = v_calculo + str((int(v_st_calculo[4]) * 1))
            v_calculo = v_calculo + str((int(v_st_calculo[5]) * 2))
            v_calculo = v_calculo + str((int(v_st_calculo[6]) * 1))
            v_calculo = v_calculo + str((int(v_st_calculo[7]) * 2))
            v_calculo = v_calculo + str((int(v_st_calculo[8]) * 1))
            v_calculo = v_calculo + str((int(v_st_calculo[9]) * 2))
            v_calculo = v_calculo + str((int(v_st_calculo[10]) * 1))
            v_calculo = v_calculo + str((int(v_st_calculo[11]) * 2))
            v_soma = 0
            for x in v_calculo:
                v_soma = v_soma + int(x)
            
            dig1 = str((int(str(v_soma)[0])+1) * 10 - v_soma)   
            if dig1 == '10': 
                dig1 = '0'
            
            if (int(insc[11]) != int(dig1)): 
                return False
            else:
                v_st_calculo = insc[0:13] 
                v_calculo = (int(v_st_calculo[0]) * 3)
                v_calculo = v_calculo + (int(v_st_calculo[1]) * 2)
                v_calculo = v_calculo + (int(v_st_calculo[2]) * 11)
                v_calculo = v_calculo + (int(v_st_calculo[3]) * 10)
                v_calculo = v_calculo + (int(v_st_calculo[4]) * 9)
                v_calculo = v_calculo + (int(v_st_calculo[5]) * 8)
                v_calculo = v_calculo + (int(v_st_calculo[6]) * 7)
                v_calculo = v_calculo + (int(v_st_calculo[7]) * 6)
                v_calculo = v_calculo + (int(v_st_calculo[8]) * 5)
                v_calculo = v_calculo + (int(v_st_calculo[9]) * 4)
                v_calculo = v_calculo + (int(v_st_calculo[10]) * 3)
                v_calculo = v_calculo + (int(v_st_calculo[11]) * 2)
                
                dig2 = 11 - (v_calculo%11) 
                if dig2 >= 10:
                    dig2 = 0
               
                if (int(insc[12]) != int(dig2)): 
                    return False

################################################
#
################################################

    elif estado == 'SP':
        
        if (len(insc) != 12):
            return False
        else:
            if inscsp[0].upper() == 'P':
                v_nr_calculo = int(insc[0])
                v_nr_calculo = v_nr_calculo + int(insc[1]) * 3
                v_nr_calculo = v_nr_calculo + int(insc[2]) * 4
                v_nr_calculo = v_nr_calculo + int(insc[3]) * 5
                v_nr_calculo = v_nr_calculo + int(insc[4]) * 6
                v_nr_calculo = v_nr_calculo + int(insc[5]) * 7
                v_nr_calculo = v_nr_calculo + int(insc[6]) * 8
                v_nr_calculo = v_nr_calculo + int(insc[7]) * 10
                v_nr_digito = (v_nr_calculo%11)
                if int(v_nr_digito) >= 10:
                    v_nr_digito = 0
                
                if int(v_nr_digito) != int(insc[9]) :
                    return False
                   
            else: 
                v_nr_calculo = int(insc[0])
                v_nr_calculo = v_nr_calculo + int(insc[1]) * 3
                v_nr_calculo = v_nr_calculo + int(insc[2]) * 4
                v_nr_calculo = v_nr_calculo + int(insc[3]) * 5
                v_nr_calculo = v_nr_calculo + int(insc[4]) * 6
                v_nr_calculo = v_nr_calculo + int(insc[5]) * 7
                v_nr_calculo = v_nr_calculo + int(insc[6]) * 8
                v_nr_calculo = v_nr_calculo + int(insc[7]) * 10
                v_nr_digito = v_nr_calculo%11
                if int(v_nr_digito) >= 10:
                    v_nr_digito = 0
                
                if int(v_nr_digito) != int(insc[8]):
                    return False
                else:
                    v_nr_calculo = int(insc[0]) * 3
                    v_nr_calculo = v_nr_calculo + int(insc[1]) * 2
                    v_nr_calculo = v_nr_calculo + int(insc[2]) * 10
                    v_nr_calculo = v_nr_calculo + int(insc[3]) * 9
                    v_nr_calculo = v_nr_calculo + int(insc[4]) * 8
                    v_nr_calculo = v_nr_calculo + int(insc[5]) * 7
                    v_nr_calculo = v_nr_calculo + int(insc[6]) * 6
                    v_nr_calculo = v_nr_calculo + int(insc[7]) * 5
                    v_nr_calculo = v_nr_calculo + int(insc[8]) * 4
                    v_nr_calculo = v_nr_calculo + int(insc[9]) * 3
                    v_nr_calculo = v_nr_calculo + int(insc[10]) * 2
                    v_nr_digito = (v_nr_calculo%11)
                    if int(v_nr_digito) >= 10:
                        v_nr_digito = 0
                    
                    if int(v_nr_digito) != int(insc[11]):
                        return False

################################################
#
################################################

    elif estado == 'RJ':
        if (len(insc) != 8):
            return False
        else:
            if int(insc) == 0:
                return False
            else:
                v_nr_calculo = int(insc[0]) * 2
                v_nr_calculo = v_nr_calculo + int(insc[1]) * 7
                v_nr_calculo = v_nr_calculo + int(insc[2]) * 6
                v_nr_calculo = v_nr_calculo + int(insc[3]) * 5
                v_nr_calculo = v_nr_calculo + int(insc[4]) * 4
                v_nr_calculo = v_nr_calculo + int(insc[5]) * 3
                v_nr_calculo = v_nr_calculo + int(insc[6]) * 2
                v_nr_digito = 11 - (v_nr_calculo%11)
                if int(v_nr_digito) >= 10:
                    v_nr_digito = 0
                
                if int(v_nr_digito) != int(insc[7]) :
                    return False

################################################
#
################################################

    elif estado == 'BA': 
        if len(insc) > 9 or len(insc) < 8:
            return False
        else:
            if len(insc) == 8:
                if insc[0] in ('0','1','2','3','4','5', '8') :
                    v_nr_calculo = int(insc[0]) * 7
                    v_nr_calculo = v_nr_calculo + int(insc[1]) * 6
                    v_nr_calculo = v_nr_calculo + int(insc[2]) * 5
                    v_nr_calculo = v_nr_calculo + int(insc[3]) * 4
                    v_nr_calculo = v_nr_calculo + int(insc[4]) * 3
                    v_nr_calculo = v_nr_calculo + int(insc[5]) * 2
                    v_nr_digito = 10 - (v_nr_calculo%10)
                    if int(v_nr_digito) >= 10:
                        v_nr_digito = 0
                    
                    if int(v_nr_digito) != int(insc[7]) :
                        return False
                    else:
                        v_nr_calculo = int(insc[0]) * 8
                        v_nr_calculo = v_nr_calculo + int(insc[1]) * 7
                        v_nr_calculo = v_nr_calculo + int(insc[2]) * 6
                        v_nr_calculo = v_nr_calculo + int(insc[3]) * 5
                        v_nr_calculo = v_nr_calculo + int(insc[4]) * 4
                        v_nr_calculo = v_nr_calculo + int(insc[5]) * 3
                        v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
                        v_nr_digito = 10 - (v_nr_calculo%10)
                        if int(v_nr_digito) >= 10 :
                            v_nr_digito = 0
                         
                        if int(v_nr_digito) != int(insc[6]) :
                            return False

                else:
                    v_nr_calculo = int(insc[0]) * 7
                    v_nr_calculo = v_nr_calculo + int(insc[1]) * 6
                    v_nr_calculo = v_nr_calculo + int(insc[2]) * 5
                    v_nr_calculo = v_nr_calculo + int(insc[3]) * 4
                    v_nr_calculo = v_nr_calculo + int(insc[4]) * 3
                    v_nr_calculo = v_nr_calculo + int(insc[5]) * 2
                    v_nr_digito = 11 - (v_nr_calculo%11)
                    if int(v_nr_digito) >= 10 :
                        v_nr_digito = 0
                     
                    if int(v_nr_digito) != int(insc[7]) :
                        return False
                    else:
                        v_nr_calculo = int(insc[0]) * 8
                        v_nr_calculo = v_nr_calculo + int(insc[1]) * 7
                        v_nr_calculo = v_nr_calculo + int(insc[2]) * 6
                        v_nr_calculo = v_nr_calculo + int(insc[3]) * 5
                        v_nr_calculo = v_nr_calculo + int(insc[4]) * 4
                        v_nr_calculo = v_nr_calculo + int(insc[5]) * 3
                        v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
                        v_nr_digito = 11 - (v_nr_calculo%11)
                        if int(v_nr_digito) >= 10 :
                            v_nr_digito = 0
                         
                        if int(v_nr_digito) != int(insc[6]) :
                            return False
                         
            else:
                if insc[1] in ('0','1','2','3','4','5','8') :
                    v_nr_calculo = int(insc[0]) * 8
                    v_nr_calculo = v_nr_calculo + int(insc[1]) * 7
                    v_nr_calculo = v_nr_calculo + int(insc[2]) * 6
                    v_nr_calculo = v_nr_calculo + int(insc[3]) * 5
                    v_nr_calculo = v_nr_calculo + int(insc[4]) * 4
                    v_nr_calculo = v_nr_calculo + int(insc[5]) * 3
                    v_nr_calculo = v_nr_calculo + int(insc[6]) * 2
                    v_nr_digito = 10 - (v_nr_calculo%10)
                    if int(v_nr_digito) >= 10 :
                        v_nr_digito = 0
                     
                    if int(v_nr_digito) != int(insc[8]) :
                        return False
                    else:
                        v_nr_calculo = int(insc[0]) * 9
                        v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
                        v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
                        v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
                        v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
                        v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
                        v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
                        v_nr_calculo = v_nr_calculo + int(insc[8]) * 2
                        v_nr_digito = 10 - (v_nr_calculo%10)
                        if int(v_nr_digito) >= 10 :
                            v_nr_digito = 0
                         
                        if int(v_nr_digito) != int(insc[7]) :
                            return False
                        
                else:
                    v_nr_calculo = int(insc[0]) * 8
                    v_nr_calculo = v_nr_calculo + int(insc[1]) * 7
                    v_nr_calculo = v_nr_calculo + int(insc[2]) * 6
                    v_nr_calculo = v_nr_calculo + int(insc[3]) * 5
                    v_nr_calculo = v_nr_calculo + int(insc[4]) * 4
                    v_nr_calculo = v_nr_calculo + int(insc[5]) * 3
                    v_nr_calculo = v_nr_calculo + int(insc[6]) * 2
                    v_nr_digito = 11 - (v_nr_calculo%11)
                    if int(v_nr_digito) >= 10 :
                        v_nr_digito = 0
                     
                    if int(v_nr_digito) != int(insc[8]) :
                        return False
                    else:
                        v_nr_calculo = int(insc[0]) * 9
                        v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
                        v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
                        v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
                        v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
                        v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
                        v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
                        v_nr_calculo = v_nr_calculo + int(insc[8]) * 2
                        v_nr_digito = 11 - (v_nr_calculo%11)
                        if int(v_nr_digito) >= 10 :
                            v_nr_digito = 0
                         
                        if int(v_nr_digito) != int(insc[7]) :
                            return False
 
################################################
#
################################################

    elif estado == 'RS' :
        if len(insc) != 10 :
            return False
        else:
            v_nr_calculo = int(insc[0]) * 2
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 9
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 8
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[7]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[8]) * 2
            v_nr_digito = 11 - (v_nr_calculo%11)
            if int(v_nr_digito) >= 10 :
                v_nr_digito = 0
             
            if int(insc[9]) != int(v_nr_digito) :
                return False

################################################
#
################################################

    elif estado in ('AM','CE','ES','PB','PI','SC','SE'):
        if len(insc) != 9 :
            return False
        else:
            v_nr_calculo = int(insc[0]) * 9
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
            v_nr_digito = 11 - (v_nr_calculo%11)
            if int(v_nr_digito) >= 10 :
                v_nr_digito = 0
             
            if int(insc[8]) != int(v_nr_digito) :
                return False

################################################
#
################################################

    elif estado == 'GO' :
        if len(insc) != 9:
            return False
        if insc[0:2] not in ('10','11','15'):
            return False
        else:
            v_nr_calculo = int(insc[0]) * 9
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
            v_nr_digito = 11 - (v_nr_calculo%11)
            if (insc[0:8] == '11094402' and ( insc[8] == '0' or insc[8] == '1' )) :
                return True
            else:    
                
                if int(v_nr_digito) == 11 :
                    v_nr_digito = 0
                 
                if int(v_nr_digito) == 10 :
                    if int(insc[0:8]) >= 10103105 and int(insc[0:8]) <= 10119997 :
                        v_nr_digito = 1
                    else:
                        v_nr_digito = 0
 
                if int(insc[8]) != int(v_nr_digito) :
                    return False
 
################################################
#
################################################

    elif estado == 'MA' :
        if len(insc) != 9 or insc[0:2] != '12' :
            return False
        else:
            v_nr_calculo = int(insc[0]) * 9
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
            v_nr_digito = 11 - (v_nr_calculo%11)
            if int(v_nr_digito) >= 10 :
                v_nr_digito = 0
             
            if int(insc[8]) != int(v_nr_digito) :
                return False

################################################
#
################################################

    elif estado == 'MS' :
        if len(insc) != 9 or insc[0:2] != '28' :
            return False
        else:
            v_nr_calculo = int(insc[0]) * 9
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
            v_nr_digito = 11 - (v_nr_calculo%11)
            if int(v_nr_digito) >= 10 :
                v_nr_digito = 0
             
            if int(insc[8]) != int(v_nr_digito) :
                return False

################################################
#
################################################
    
    elif estado == 'PA' :
        if len(insc) != 9 or insc[0:2] != '15' :
            return False
        else:
            v_nr_calculo = int(insc[0]) * 9
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
            v_nr_digito = 11 - (v_nr_calculo%11)
            if int(v_nr_digito) >= 10 :
                v_nr_digito = 0
             
            if int(insc[8]) != int(v_nr_digito) :
                return False
 
################################################
#
################################################

    elif estado == 'DF':
        if len(insc) != 13 or insc[0:2] != '07' :
            return False
        else:
            v_nr_calculo = int(insc[0]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 2
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 9
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 8
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[7]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[8]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[9]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[10]) * 2
            v_nr_digito = 11 - (v_nr_calculo%11)
            if int(v_nr_digito) >= 10 :
                v_nr_digito = 0
             
            if int(insc[11]) != int(v_nr_digito) :
                return False
            else:
                v_nr_calculo = int(insc[0]) * 5
                v_nr_calculo = v_nr_calculo + int(insc[1]) * 4
                v_nr_calculo = v_nr_calculo + int(insc[2]) * 3
                v_nr_calculo = v_nr_calculo + int(insc[3]) * 2
                v_nr_calculo = v_nr_calculo + int(insc[4]) * 9
                v_nr_calculo = v_nr_calculo + int(insc[5]) * 8
                v_nr_calculo = v_nr_calculo + int(insc[6]) * 7
                v_nr_calculo = v_nr_calculo + int(insc[7]) * 6
                v_nr_calculo = v_nr_calculo + int(insc[8]) * 5
                v_nr_calculo = v_nr_calculo + int(insc[9]) * 4
                v_nr_calculo = v_nr_calculo + int(insc[10]) * 3
                v_nr_calculo = v_nr_calculo + int(insc[11]) * 2
                v_nr_digito = 11 - (v_nr_calculo%11)
                if int(v_nr_digito) >= 10 :
                    v_nr_digito = 0
                 
                if int(insc[12]) != int(v_nr_digito) :
                    return False

################################################
#
################################################

    elif estado == 'TO' :
        if str(len(insc)) not in ('9','11'):
            return False
        else:
            
            if len(insc) == 11 :
                if insc[2:4] not in ('01','02','03','99'):
                    return False
                else:
                    v_nr_calculo = int(insc[0]) * 9
                    v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
                    v_nr_calculo = v_nr_calculo + int(insc[4]) * 7
                    v_nr_calculo = v_nr_calculo + int(insc[5]) * 6
                    v_nr_calculo = v_nr_calculo + int(insc[6]) * 5
                    v_nr_calculo = v_nr_calculo + int(insc[7]) * 4
                    v_nr_calculo = v_nr_calculo + int(insc[8]) * 3
                    v_nr_calculo = v_nr_calculo + int(insc[9]) * 2
                    if (v_nr_calculo%11) < 2 :
                        v_nr_digito = 0
                     
                    if (v_nr_calculo%11) >= 2 :
                        v_nr_digito = 11 - (v_nr_calculo%11)
   
                    if int(insc[10]) != int(v_nr_digito) :
                        return False

            if len(insc) == 9 :
                v_nr_calculo = int(insc[0]) * 9
                v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
                v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
                v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
                v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
                v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
                v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
                v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
                if (v_nr_calculo%11) < 2 :
                    v_nr_digito = 0
                 
                if (v_nr_calculo%11) >= 2 :
                    v_nr_digito = 11 - (v_nr_calculo%11)

                if int(insc[8]) != int(v_nr_digito) :
                    return False

################################################
#
################################################

    elif estado == 'PR' :
        if len(insc) != 10 :
            return False
        else:
            v_nr_calculo = int(insc[0]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 2
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
            v_nr_digito = 11 - (v_nr_calculo%11)
            if int(v_nr_digito) >= 10 :
                v_nr_digito = 0
             
            if int(insc[8]) != int(v_nr_digito) :
                return False
            else:
                v_nr_calculo = int(insc[0]) * 4
                v_nr_calculo = v_nr_calculo + int(insc[1]) * 3
                v_nr_calculo = v_nr_calculo + int(insc[2]) * 2
                v_nr_calculo = v_nr_calculo + int(insc[3]) * 7
                v_nr_calculo = v_nr_calculo + int(insc[4]) * 6
                v_nr_calculo = v_nr_calculo + int(insc[5]) * 5
                v_nr_calculo = v_nr_calculo + int(insc[6]) * 4
                v_nr_calculo = v_nr_calculo + int(insc[7]) * 3
                v_nr_calculo = v_nr_calculo + int(insc[8]) * 2
                v_nr_digito = 11 - (v_nr_calculo%11)
                if int(v_nr_digito) >= 10 :
                    v_nr_digito = 0
                 
                if int(insc[9]) != int(v_nr_digito) :
                    return False
 
################################################
#
################################################

    elif estado == 'RO' :
        if len(insc) != 14 :
            return False
        else:
            v_nr_calculo = int(insc[0]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 2
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 9
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 8
            v_nr_calculo = v_nr_calculo + int(insc[7]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[8]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[9]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[10]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[11]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[12]) * 2
            v_nr_digito = 11 - (v_nr_calculo%11)
            if int(v_nr_digito) == 10 :
                v_nr_digito = 0
             
            if int(v_nr_digito) == 11 :
                v_nr_digito = 1
             
            if int(insc[13]) != int(v_nr_digito) :
                return False
 
################################################
#
################################################

    elif estado == 'RR' :
        if len(insc) != 9 or insc[0:2] != '24' :
            return False
        else:
            v_nr_calculo = int(insc[0]) * 1
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 2
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[7]) * 8
            v_nr_digito = (v_nr_calculo%9)
            
            if int(insc[8]) != int(v_nr_digito) :
                return False

################################################
#
################################################

    elif estado == 'PE' :
        if len(insc) != 9 :
            return False
        else:
            v_nr_calculo = int(insc[0]) * 8
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 2
            v_nr_digito = 11 - (v_nr_calculo%11)
            if int(v_nr_digito) >= 10 :
                v_nr_digito = 0
             
            if int(insc[7]) != int(v_nr_digito) :
                return False
            else:
                v_nr_calculo = int(insc[0]) * 9
                v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
                v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
                v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
                v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
                v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
                v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
                v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
                v_nr_digito = 11 - (v_nr_calculo%11)
                if int(v_nr_digito) >= 10 :
                    v_nr_digito = 0
                 
                if int(insc[8]) != int(v_nr_digito) :
                    return False

################################################
#
################################################

    elif estado == 'RN' :
        if insc[0:2] != '20' :
            return False
        else:
            if len(insc) == 9 :
                v_nr_calculo = int(insc[0]) * 9
                v_nr_calculo = v_nr_calculo + int(insc[1]) * 8
                v_nr_calculo = v_nr_calculo + int(insc[2]) * 7
                v_nr_calculo = v_nr_calculo + int(insc[3]) * 6
                v_nr_calculo = v_nr_calculo + int(insc[4]) * 5
                v_nr_calculo = v_nr_calculo + int(insc[5]) * 4
                v_nr_calculo = v_nr_calculo + int(insc[6]) * 3
                v_nr_calculo = v_nr_calculo + int(insc[7]) * 2
                v_nr_calculo = v_nr_calculo * 10
                v_nr_digito = (v_nr_calculo%11)
                if int(v_nr_digito) >= 10 :
                    v_nr_digito = 0
                 
                if int(insc[8]) != int(v_nr_digito) :
                    return False
                 

            else:
                if len(insc) != 10 :
                    return False
                else:
                    v_nr_calculo = int(insc[0]) * 10
                    v_nr_calculo = v_nr_calculo + int(insc[1]) * 9
                    v_nr_calculo = v_nr_calculo + int(insc[2]) * 8
                    v_nr_calculo = v_nr_calculo + int(insc[3]) * 7
                    v_nr_calculo = v_nr_calculo + int(insc[4]) * 6
                    v_nr_calculo = v_nr_calculo + int(insc[5]) * 5
                    v_nr_calculo = v_nr_calculo + int(insc[6]) * 4
                    v_nr_calculo = v_nr_calculo + int(insc[7]) * 3
                    v_nr_calculo = v_nr_calculo + int(insc[8]) * 2
                    v_nr_calculo = v_nr_calculo * 10
                    v_nr_digito = (v_nr_calculo%11)
                    if int(v_nr_digito) >= 10 :
                        v_nr_digito = 0
                     
                    if int(insc[9]) != int(v_nr_digito) :
                        return False
 
################################################
#
################################################

    elif estado == 'MT' :
        if len(insc) != 11 :
            return False
        else:
            v_nr_calculo = int(insc[0]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[1]) * 2
            v_nr_calculo = v_nr_calculo + int(insc[2]) * 9
            v_nr_calculo = v_nr_calculo + int(insc[3]) * 8
            v_nr_calculo = v_nr_calculo + int(insc[4]) * 7
            v_nr_calculo = v_nr_calculo + int(insc[5]) * 6
            v_nr_calculo = v_nr_calculo + int(insc[6]) * 5
            v_nr_calculo = v_nr_calculo + int(insc[7]) * 4
            v_nr_calculo = v_nr_calculo + int(insc[8]) * 3
            v_nr_calculo = v_nr_calculo + int(insc[9]) * 2
            v_nr_digito = 11 - (v_nr_calculo%11)
            if int(v_nr_digito) >= 10 :
                v_nr_digito = 0
             
            if int(insc[10]) != int(v_nr_digito) :
                return False

################################################
#
################################################
    else:
        return False   
    
    return True

