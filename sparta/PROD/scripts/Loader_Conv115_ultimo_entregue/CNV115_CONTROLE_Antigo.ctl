--   CARGA MANUAL SIA/NSIA - ITEM      

options (errors=9999999, rows=1000)
LOAD DATA      

    infile '$$IN$$'

INTO TABLE gfcarga.tsh_controle_conv_115_ent
Append      
(      
    id_serie_levantamento              " '$$ID$$' "         ,
    uf_filial                          " '$$UF$$' "         ,
    volume                             " '$$VOL$$' "        ,
    serie                              " '$$SERIE$$' "      ,

    CNPJ                               position(1:18)       ,
    IE                                 position(19:33)      ,
    RAZAO_SOCIAL                       position(34:83)      ,
    ENDERECO                           position(84:133)     ,
    CEP                                position(134:142)    ,
    BAIRRO                             position(143:172)    ,
    MUNICIPIO                          position(173:202)    ,
    UF                                 position(203:204)    ,
    RESPONSAVEL_APRESENTACAO           position(205:234)    ,
    CARGO                              position(235:254)    ,
    TELEFONE                           position(255:266)    ,
    E_MAIL                             position(267:306)    ,
    QTD_REGISTRO_NF                    position(307:313)    ,
    QTD_REGISTRO_CANCELADO_NF          position(314:320)    ,
    DATA_EMISSAO_PRIM_DOC_NF           position(321:328)    "to_date(:DATA_EMISSAO_PRIM_DOC_NF, 'YYYYMMDD')",
    DATA_EMISSAO_ULT_DOC_NF            position(329:336)    "to_date(:DATA_EMISSAO_ULT_DOC_NF, 'YYYYMMDD')",
    NUMERO_PRIMEIRO_NF                 position(337:345)    ,
    NUMERO_ULTIMO_NF                   position(346:354)    ,
    VALOR_TOTAL_NF                     position(355:368)    ZONED(14,2),
    BC_ICMS_NF                         position(369:382)    ZONED(14,2),
    ICMS_NF                            position(383:396)    ZONED(14,2),
    ISENTAS_E_NAO_TRIB_NF              position(397:410)    ZONED(14,2),
    OUTROS_NF                          position(411:424)    ZONED(14,2),
    NOMENCLATURA_ARQ_NF                position(425:439)    ,
    STATUS_RETIFICACAO_NF              position(440:440)    ,
    HASHCOD_NF                         position(441:472)    ,
    QTD_REGISTRO_ITEM                  position(473:481)    ,
    QTD_REGISTRO_CANCELADO_ITEM        position(482:488)    ,
    DATA_EMISSAO_PRIM_DOC_ITEM         position(489:496)    "to_date(:DATA_EMISSAO_PRIM_DOC_ITEM, 'YYYYMMDD')",
    DATA_EMISSAO_ULT_DOC_ITEM          position(497:504)    "to_date(:DATA_EMISSAO_ULT_DOC_ITEM, 'YYYYMMDD')",
    NUMERO_PRIMEIRO_ITEM               position(505:513)    ,
    NUMERO_ULTIMO_ITEM                 position(514:522)    ,
    VALOR_TOTAL_ITEM                   position(523:536)    ZONED(14,2),
    DESCONTOS_ITEM                     position(537:550)    ZONED(14,2),
    ACRESCIMO_DESP_ACESSORIAS_ITEM     position(551:564)    ZONED(14,2),
    BC_ICMS_ITEM                       position(565:578)    ZONED(14,2),
    ICMS_ITEM                          position(579:592)    ZONED(14,2),
    ISENTAS_E_NAO_TRIB_ITEM            position(593:606)    ZONED(14,2),
    OUTROS_ITEM                        position(607:620)    ZONED(14,2),
    NOMENCLATURA_ARQ_ITEM              position(621:635)    ,
    STATUS_RETIFICACAO_ITEM            position(636:636)    ,
    HASHCOD_ITEM                       position(637:668)    ,
    QTD_REGISTRO_DESTINATARIO          position(669:675)    ,
    NOMENCLATURA_ARQ_DESTINATARIO      position(676:690)    ,
    ST_RETIFICACAO_DESTINATARIO        position(691:691)    ,
    HASHCOD_DESTINATARIO               position(692:723)    ,
    VERSAO_PROG                        position(724:726)    ,
    CHV_CONTROLE_REC_ENTREGA           position(727:732)    ,
    QTD_ADVERTENCIAS                   position(733:741)    ,
    HASHCOD_REGISTRO                   position(766:797)     

)
