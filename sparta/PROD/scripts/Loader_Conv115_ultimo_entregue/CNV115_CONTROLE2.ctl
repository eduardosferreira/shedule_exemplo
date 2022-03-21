--   CARGA MANUAL SIA/NSIA - ITEM      

options (errors=9999999, rows=1000)
LOAD DATA      

infile '$$IN$$'

INTO TABLE gfcarga.tsh_controle_conv_115_ent
Append      
WHEN VOLUME != "X"
(      
    id_serie_levantamento            	" '$$ID$$' "       	 ,
    uf_filial                        	" '$$UF$$' "       	 ,
	
    CNPJ                             	position(1:18)       ,
    IE                               	position(19:33)      ,
    RAZAO_SOCIAL                     	position(34:83)      ,
    ENDERECO                         	position(84:133)     ,
    CEP                              	position(134:142)    ,
    BAIRRO                           	position(143:172)    ,
    MUNICIPIO                        	position(173:202)    ,
    UF                               	position(203:204)    ,
    RESPONSAVEL_APRESENTACAO         	position(205:234)    ,
    CARGO                            	position(235:254)    ,
    TELEFONE                         	position(255:266)    ,
    E_MAIL                           	position(267:306)    ,
    QTD_REGISTRO_NF                  	position(307:313)    ,
    QTD_REGISTRO_CANCELADO_NF        	position(314:320)    ,
    DATA_EMISSAO_PRIM_DOC_NF         	position(321:328)    "to_date(:DATA_EMISSAO_PRIM_DOC_NF, 'YYYYMMDD')",
    DATA_EMISSAO_ULT_DOC_NF          	position(329:336)    "to_date(:DATA_EMISSAO_ULT_DOC_NF, 'YYYYMMDD')",
    NUMERO_PRIMEIRO_NF               	position(337:345)    ,
    NUMERO_ULTIMO_NF                 	position(346:354)    ,
    VALOR_TOTAL_NF                   	position(355:368)    ZONED(14,2),
    BC_ICMS_NF                       	position(369:382)    ZONED(14,2),
    ICMS_NF                          	position(383:396)    ZONED(14,2),
    ISENTAS_E_NAO_TRIB_NF            	position(397:410)    ZONED(14,2),
    OUTROS_NF                        	position(411:424)    ZONED(14,2),
    NOMENCLATURA_ARQ_NF              	position(425:464)    ,
    STATUS_RETIFICACAO_NF            	position(465:465)    ,
    HASHCOD_NF                       	position(466:497)    ,
    QTD_REGISTRO_ITEM                	position(498:506)    ,
    QTD_REGISTRO_CANCELADO_ITEM      	position(507:513)    ,
    DATA_EMISSAO_PRIM_DOC_ITEM       	position(514:521)    "to_date(:DATA_EMISSAO_PRIM_DOC_ITEM, 'YYYYMMDD')",
    DATA_EMISSAO_ULT_DOC_ITEM        	position(522:529)    "to_date(:DATA_EMISSAO_ULT_DOC_ITEM, 'YYYYMMDD')",
    NUMERO_PRIMEIRO_ITEM             	position(530:538)    ,
    NUMERO_ULTIMO_ITEM               	position(539:547)    ,
    VALOR_TOTAL_ITEM                 	position(548:561)    ZONED(14,2),
    DESCONTOS_ITEM                   	position(562:575)    ZONED(14,2),
    ACRESCIMO_DESP_ACESSORIAS_ITEM   	position(576:589)    ZONED(14,2),
    BC_ICMS_ITEM                     	position(590:603)    ZONED(14,2),
    ICMS_ITEM                        	position(604:617)    ZONED(14,2),
    ISENTAS_E_NAO_TRIB_ITEM          	position(618:631)    ZONED(14,2),
    OUTROS_ITEM                      	position(632:645)    ZONED(14,2),
    NOMENCLATURA_ARQ_ITEM            	position(646:685)    ,
    STATUS_RETIFICACAO_ITEM          	position(686:686)    ,
    HASHCOD_ITEM                     	position(687:718)    ,
    QTD_REGISTRO_DESTINATARIO        	position(719:725)    ,
    NOMENCLATURA_ARQ_DESTINATARIO    	position(726:765)    ,
    ST_RETIFICACAO_DESTINATARIO      	position(766:766)    ,
    HASHCOD_DESTINATARIO             	position(767:798)    ,
    VERSAO_PROG                      	position(799:801)    ,
    CHV_CONTROLE_REC_ENTREGA         	position(802:807)    ,
    QTD_ADVERTENCIAS                 	position(808:816)    ,
    REF_APURACAO                     	position(817:820)    ,
    MODELO                           	position(821:822)    ,
    SERIE                            	position(823:825)    ,
    VOLUME                           	position(826:828)    "NVL(:VOLUME,'X')",
    SITUACAO_VERSAO                  	position(829:831)    ,
    NOMENCLATURA_ARQ_COMPACTADO      	position(832:891)    ,
    HASHCOD_REGISTRO                 	position(1304:1335)  

)      

