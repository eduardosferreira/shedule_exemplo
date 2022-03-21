--   CARGA MANUAL SIA/NSIA - MESTRE      

options (errors=9999999, rows=1000)
Load DATA
  infile '/portaloptrib/LEVCV115/RJ/16/01/TBRA/1710/SERIE/96005938/PROTOCOLADO/RJ11 1601NM.001'

INTO TABLE SPT80732427.tsh_mestre_conv_115
Append      
(      

id_serie_levantamento   " '96005938' "        ,
uf_filial               " 'RJ' "        ,
volume                  " '001' "       ,
linha                   RECNUM              ,

CNPJ_CPF                POSITION(1:14)     "TRIM(:CNPJ_CPF)",
IE                      POSITION(15:28)    "TRIM(:IE)",
RAZAO_SOCIAL            POSITION(29:63)    "TRIM(:RAZAO_SOCIAL)",
UF                      POSITION(64:65)     ,
CLASSE_CONS             POSITION(66:66)     ,
TIPO_UTILIZ             POSITION(67:67)     ,
GRUPO_TENSAO            POSITION(68:69)     ,
CADG_COD                POSITION(70:81)     ,
DATA_EMISSAO            POSITION(82:89)     "to_date(:DATA_EMISSAO, 'YYYYMMDD') ",
MODELO                  POSITION(90:91)     ,
SERIE                   POSITION(92:94)     ,
NUMERO_NF               POSITION(95:103)    ,
HASH_COD_NF             POSITION(104:135)   , -- Acrescentado
VALOR_TOTAL             POSITION(136:147)   ZONED(12,2),
BASE_ICMS               POSITION(148:159)   ZONED(12,2),
VALOR_ICMS              POSITION(160:171)   ZONED(12,2),
ISENTAS_ICMS            POSITION(172:183)   ZONED(12,2),
OUTROS_VALORES          POSITION(184:195)   ZONED(12,2),
SIT_DOC                 POSITION(196:196)   ,
MES_REFERENCIA          POSITION(197:200)   ,
REF_ITEM_NF             POSITION(201:209)   ,
TERMINAL_TELEF          POSITION(210:221)   "TRIM(:TERMINAL_TELEF)",
IND_CAMPO_01            POSITION(222:222)   ,
TIPO_CLIENTE            POSITION(223:224)   ,
SUB_CLASSE              POSITION(225:226)   ,
TERMINAL_PRINC          " '' "              ,
CNPJ_EMIT               " '' "              ,
NUM_FAT                 " '' "              ,
VALOR_FAT               " '' "              ,
BRANCOS2                " '' "              ,
HASH_CODE_ARQ           POSITION(227:258)   

)
