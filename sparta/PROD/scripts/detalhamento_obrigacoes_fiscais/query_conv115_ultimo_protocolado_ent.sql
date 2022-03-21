SELECT /*+ parallel(8)*/
     L.EMPS_COD                      "Empresa"
    ,'[uf_filial]'                   "UF Filial"
    ,L.FILI_COD                      "Filial"
    ,'[mes_ano]'                     "Mês / Ano"
    ,trim(I.SERIE)                   "Série"
    ,I.ALIQUOTA                      "Alíquota"
    ,I.CFOP                          "CFOP"
    ,I.VOLUME                        "Volume"
    ,SUM(I.VALOR_TOTAL - I.DESCONTO) "Valor Líquido"
    ,SUM(I.BASE_ICMS )               "Valor Base"
    ,SUM(I.VALOR_ICMS )              "Valor de ICMS"
    ,SUM(I.ISENTAS_ICMS )            "Valor de Isentas"
    ,SUM(I.OUTROS_VALORES )          "Valor de Outras"
    ,SUM(I.VALOR_TOTAL )             "Valor Total"
    ,SUM(I.DESCONTO )                "Desconto / Redutores"
    ,C.NOMENCLATURA_ARQ_NF           "Nome do Arquivo Mestre"
    ,C.HASHCOD_NF		     "Código de Autenticação"
    ,c.STATUS_RETIFICACAO_NF	     "Indicador Retificação"	 	
  FROM gfcarga.TSH_ITEM_CONV_115_ENT[C6]  I
  JOIN gfcarga.TSH_SERIE_LEVANTAMENTO[C6] L
    ON L.id_serie_levantamento  = I.id_serie_levantamento
   AND L.uf_filial              = I.uf_filial
  JOIN gfcarga.TSH_CONTROLE_CONV_115_ENT[C6] C
    ON I.id_serie_levantamento  = C.id_serie_levantamento
   AND I.uf_filial              = C.uf_filial
   AND TO_NUMBER(I.NUMERO_NF) BETWEEN C.NUMERO_PRIMEIRO_NF AND C.NUMERO_ULTIMO_NF
 WHERE L.uf_filial              = '[uf_filial]'
   AND l.mes_ano                = to_date('[datai]', 'DD/MM/YYYY')
   AND l.ORIGEM_PROTOCOLADO     = 'ULTIMO_ENTREGUE'
   AND I.sit_doc   		= 'N'
 GROUP BY L.EMPS_COD, L.FILI_COD, trim(I.SERIE), I.ALIQUOTA,I.CFOP,I.VOLUME,C.NOMENCLATURA_ARQ_NF,C.HASHCOD_NF,c.STATUS_RETIFICACAO_NF
