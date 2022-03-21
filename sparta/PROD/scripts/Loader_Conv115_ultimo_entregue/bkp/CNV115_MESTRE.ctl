--   CARGA MANUAL SIA/NSIA - MESTRE      

options (errors=9999999, rows=1000)
Load DATA
  infile '$$IN$$'

INTO TABLE gfcarga.tsh_mestre_conv_115_ent
Append      
(      

id_serie_levantamento   " '$$ID$$' "        ,
uf_filial               " '$$UF$$' "        ,
volume                  " '$$VOL$$' "       ,
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
TERMINAL_PRINC          POSITION(227:238)   ,
CNPJ_EMIT               POSITION(239:252)   ,
NUM_FAT                 POSITION(253:272)   ,
VALOR_FAT               POSITION(273:284)   ZONED(12,2),
BRANCOS2                POSITION(285:393)   ,
HASH_CODE_ARQ           POSITION(394:425)   

)
