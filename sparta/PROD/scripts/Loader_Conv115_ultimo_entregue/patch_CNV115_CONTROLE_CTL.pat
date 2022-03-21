--- CNV115_CONTROLE2.ctl	2022-03-18 12:36:54.008487000 -0300
+++ CNV115_CONTROLE_JUST.ctl	2022-03-18 12:36:29.085895000 -0300
@@ -7,6 +7,7 @@
 
 INTO TABLE gfcarga.tsh_controle_conv_115_ent
 Append      
+WHEN VOLUME != "X"
 (      
     id_serie_levantamento            	" '$$ID$$' "       	 ,
     uf_filial                        	" '$$UF$$' "       	 ,
@@ -63,7 +64,7 @@
     REF_APURACAO                     	position(817:820)    ,
     MODELO                           	position(821:822)    ,
     SERIE                            	position(823:825)    ,
-    VOLUME                           	position(826:828)    ,
+    VOLUME                           	position(826:828)    "NVL(:VOLUME,'X')",
     SITUACAO_VERSAO                  	position(829:831)    ,
     NOMENCLATURA_ARQ_COMPACTADO      	position(832:891)    ,
     HASHCOD_REGISTRO                 	position(1304:1335)  
