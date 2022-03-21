--   CARGA MANUAL SIA/NSIA - ITEM      

options (errors=9999999, rows=1000)
LOAD DATA      

infile '$$IN$$'

INTO TABLE gfcarga.tsh_item_conv_115
Append      
(      
    id_serie_levantamento    " '$$ID$$' "           ,
    uf_filial                " '$$UF$$' "           ,
    volume                   " '$$VOL$$' "          ,
    linha                    RECNUM                 ,

    cnpj_cpf                 POSITION(1:14)         "TRIM(:CNPJ_CPF)",
    uf                       POSITION(15:16)        ,
    classe_cons              POSITION(17:17)        ,
    tipo_utiliz              POSITION(18:18)        ,
    data_emissao             POSITION(21:28)        "to_date(:data_emissao, 'YYYYMMDD')",
    modelo                   POSITION(29:30)        ,
    serie                    POSITION(31:33)        ,
    numero_nf                POSITION(34:42)        ,
    cfop                     POSITION(43:46)        ,
    num_item                 POSITION(47:49)        ,
    cod_item                 POSITION(50:59)        ,
    descr_item               POSITION(60:99)        "TRIM(:descr_item)",
    class_item               POSITION(100:103)      ,
    unidade                  POSITION(104:109)      "TRIM(:UNIDADE)",
    quantidade_contr         POSITION(110:121)      ZONED(12,3),
    quantidade_med           POSITION(122:133)      ZONED(12,3),
    valor_total              POSITION(134:144)      ZONED(11,2),
    desconto                 POSITION(145:155)      ZONED(11,2),
    acresc_desp              POSITION(156:166)      ZONED(11,2),
    base_icms                POSITION(167:177)      ZONED(11,2),
    valor_icms               POSITION(178:188)      ZONED(11,2),
    isentas_icms             POSITION(189:199)      ZONED(11,2),
    outros_valores           POSITION(200:210)      ZONED(11,2),
    aliquota                 POSITION(211:214)      ZONED(4,2),
    sit_doc                  POSITION(215:215)      ,
    mes_referencia           POSITION(216:219)      ,
    num_contr                POSITION(220:234)      ,
    quant_faturada           POSITION(235:246)      ZONED(12,3),
    aliq_pis                 POSITION(258:263)      ZONED(6,4),
    vlr_pis                  POSITION(264:274)      ZONED(11,2),
    aliq_cofins              POSITION(275:280)      ZONED(6,4),
    vlr_cofins               POSITION(281:291)      ZONED(11,2),
    ind_desc_jud             POSITION(292:292)      ,
    tipo_isencao             POSITION(293:294)      ,
    hash_code_arq            POSITION(300:331)

)      
