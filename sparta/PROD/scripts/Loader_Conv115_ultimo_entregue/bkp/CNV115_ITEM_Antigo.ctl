--   CARGA MANUAL SIA/NSIA - ITEM      

options (errors=9999999, rows=1000)
LOAD DATA      

    infile '$$IN$$'

INTO TABLE gfcarga.tsh_item_conv_115_ent
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
    quantidade_contr         POSITION(110:120)      ZONED(11,3),
    quantidade_med           POSITION(121:131)      ZONED(11,3),
    valor_total              POSITION(132:142)      ZONED(11,2),
    desconto                 POSITION(143:153)      ZONED(11,2),
    acresc_desp              POSITION(154:164)      ZONED(11,2),
    base_icms                POSITION(165:175)      ZONED(11,2),
    valor_icms               POSITION(176:186)      ZONED(11,2),
    isentas_icms             POSITION(187:197)      ZONED(11,2),
    outros_valores           POSITION(198:208)      ZONED(11,2),
    aliquota                 POSITION(209:212)      ZONED(4,2),
    sit_doc                  POSITION(213:213)      ,
    mes_referencia           POSITION(214:217)      ,
    num_contr                " '' "                 ,
    quant_faturada           " '0.00' "             ,
    aliq_pis                 " '0.00' "             ,
    vlr_pis                  " '0.00' "             ,
    aliq_cofins              " '0.00' "             ,
    vlr_cofins               " '0.00' "             ,
    ind_desc_jud             " '' "                 ,
    tipo_isencao             " '' "                 ,
    hash_code_arq            POSITION(223:254)

)      
