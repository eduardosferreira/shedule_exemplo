OPTIONS (SKIP=1, DIRECT=FALSE, PARALLEL=TRUE, ERRORS=1000000000, bindsize=100000, ROWS=10000)

LOAD DATA
infile '$$IN$$'

INTO TABLE gfcarga.TSH_CONV_86
APPEND
WHEN (1:1) = '2'
TRAILING NULLCOLS
(
   MODELO_NOTA_FISCAL          POSITION(2:3),
   NUMERO_NOTA_FISCAL          POSITION(4:12),
   SERIE                       POSITION(13:15),
   DATA_EMISSAO                POSITION(16:23)   DATE 'YYYYMMDD',
   CODIGO_AUTENTICACAO_DIGITAL POSITION(24:55),
   CNPJ_CPF                    POSITION(56:69),
   INSCRICAO_ESTADUAL          POSITION(70:83),
   RAZAO_COCIAL                POSITION(84:118),
   CODIGO_CONSUMIDOR           POSITION(119:130),
   VALOR_TOTAL_NOTA_FISCAL     POSITION(131:142) ZONED(12,2),
   BASE_ICMS_NOTA_FISCAL       POSITION(143:154) ZONED(12,2),
   VALOR_ICMS_NOTA_FISCAL      POSITION(155:166) ZONED(12,2),
   NUMERO_ITEM                 POSITION(167:169) ZONED(3),
   VALOR_ITEM                  POSITION(170:181) ZONED(12,2),
   VALOR_ESTORNO               POSITION(182:193) ZONED(12,2),
   HIPOTESE_ESTORNO            POSITION(194:194),
   MOTIVO                      POSITION(195:394),
   NUMERO_RECLAMACAO           POSITION(395:414),
   NOME_ARQUIVO                CONSTANT '$$ARQ$$',
   UF_FILIAL                   CONSTANT '$$UF$$',
   EMPS_COD                    CONSTANT 'TBRA',
   VERSAO                      CONSTANT '$$VER$$'
)
