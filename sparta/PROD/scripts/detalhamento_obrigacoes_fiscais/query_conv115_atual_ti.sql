with controle as (
    select 
          "Volume"
        , "Nome do Arquivo Mestre"
        , "Código de Autenticação"
        , "Indicador Retificação"
        , CTR_NUM_NFINI
        , CTR_NUM_NFFIN
        , CTR_APUR_DTINI
        , CTR_SERIE
        , FILI_COD
        , EMPS_COD
    from (
    select /*+ parallel(20)*/
      c.EMPS_COD
    , l.uf_filial 
    , c.FILI_COD
    , l.mes_ano 
    , c.CTR_SERIE
    , to_number(substr(c.CTR_NF_NOMARQ,-3,3)) "Volume"
    , c.CTR_NF_NOMARQ   "Nome do Arquivo Mestre"
    , c.CTR_CODH_ARQNF  "Código de Autenticação"
    , c.CTR_IND_RETIF   "Indicador Retificação"
    , c.CTR_NUM_NFINI
    , c.CTR_NUM_NFFIN 
    , c.CTR_APUR_DTINI
    , RANK() OVER (PARTITION BY c.EMPS_COD, c.FILI_COD, c.CTR_SERIE, TO_NUMBER(c.CTR_VOLUME) ORDER BY CTR_IND_RETIF desc) SEQ_RETIFICACAO
    from openrisow.CTR_IDENT_CNV115 c
    join gfcarga.TSH_SERIE_LEVANTAMENTO l
    on c.emps_cod = l.emps_cod
    and c.fili_cod = l.fili_cod
    and to_char(c.CTR_APUR_DTINI,'MMYYYY') = to_char(l.mes_ano,'MMYYYY')
    and c.CTR_SERIE = replace(l.serie,' ','')
    where l.mes_ano =  to_date('[datai]', 'DD/MM/YYYY')
    and l.uf_filial = '[uf_filial]'
    )
    where SEQ_RETIFICACAO = 1
    order by EMPS_COD, uf_filial, CTR_SERIE
) SELECT /*+ parallel(20)*/
     L.EMPS_COD "Empresa"
    ,L.UF_FILIAL "UF Filial"
    ,L.FILI_COD "Filial"
    ,to_char(L.MES_ANO, 'MMYYYY') "Mês / Ano"
    ,replace(I.INFST_SERIE,' ','') "Série"
    ,"Volume"
    ,decode(I.infst_tribicms,'S',I.infst_aliq_icms,0) "Alíquota"
    ,'0' || I.ESTB_COD         "CST"
    ,I.CFOP 		       "CFOP"
    ,SUM(I.INFST_VAL_CONT)     "Valor Líquido"
    ,SUM(I.INFST_BASE_ICMS )   "Valor Base"
    ,SUM(I.INFST_VAL_ICMS )    "Valor de ICMS"
    ,SUM(I.INFST_ISENTA_ICMS ) "Valor de Isentas"
    ,SUM(I.INFST_OUTRAS_ICMS ) "Valor de Outras"
    ,SUM(I.INFST_VAL_SERV )    "Valor Total"
    ,SUM(I.INFST_VAL_DESC )    "Desconto / Redutores"
    ,"Nome do Arquivo Mestre"
    ,"Código de Autenticação" 
    ,"Indicador Retificação"  
    from openrisow.ITEM_NFTL_SERV I
    join gfcarga.TSH_SERIE_LEVANTAMENTO L
    on I.emps_cod = L.emps_cod
    and I.fili_cod = L.fili_cod
    and I.infst_serie = L.serie
    and I.infst_dtemiss >= L.mes_ano
    and I.infst_dtemiss <= last_day(L.mes_ano)
    
    join controle c 
        on  c.emps_cod       = l.emps_cod
       and  c.FILI_COD       = l.fili_cod
       and trunc(CTR_APUR_DTINI,'MM') = l.mes_ano
       and c.CTR_SERIE       = replace(l.serie,' ','')
       and to_number(i.infst_num) between to_number(CTR_NUM_NFINI) and to_number(CTR_NUM_NFFIN)

    where L.uf_filial = '[uf_filial]'
    and L.mes_ano = to_date('[datai]', 'DD/MM/YYYY')
    AND i.INFST_IND_CANC  =  'N'
    group by L.EMPS_COD,L.UF_FILIAL,L.FILI_COD,L.MES_ANO,'0' || I.ESTB_COD,
             replace(I.INFST_SERIE,' ',''),decode(I.infst_tribicms,'S',I.infst_aliq_icms,0), 
             I.CFOP,"Nome do Arquivo Mestre","Volume","Código de Autenticação","Indicador Retificação"
