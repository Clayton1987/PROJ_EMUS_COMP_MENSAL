#### Lendo Excel comparando informações
### data criação: 22/12/2023
### modificado: 22/12/2023
import pandas as pd
import time
import datetime


###C:\Users\CLAY\Documents\AUTO_MAG\ANALISE\ARQ
###C:\Users\CLAY\Documents\EMUSA\MATRIZ_XML_NF_2019_2023\MATRIZ_XML_NF_2019_2023
###C:\Users\CLAY\PycharmProjects\Proj_inicial\ARQ

###LENDO ARQUIVO CADASTRO DE ITEM MENSAL

df1 = pd.read_excel(r'C:\Users\CLAY\PycharmProjects\Proj_inicial\ARQ\ANALISE\TODOS_PRODUTOS.xlsx', sheet_name='EFAPRODU')
df_agrelli = df1.loc[(df1['DTCUST'].dt.year == 2023) & (df1['DTCUST'].dt.month >= 11),['CODPRO','NOMPRO','CLAFIS','CEST','ALQICM','CODTRI','CODTR2','REDAGR','DTCUST','SALEST','SALDEP','REFERE','CDPRFO','TABFIX']].reset_index(drop=True)


#### LENDO ARQUIVO XMLS DE NFS DO MES
df2 = pd.read_excel(r'C:\Users\CLAY\PycharmProjects\Proj_inicial\ARQ\ANALISE\NFs_xml.xlsx')
df_xml = df2.drop(columns='Unnamed: 0')
df_xml_comp = df_xml[['cod_prod_xml', 'nome_prod_xml', 'ncm_prod_xml', 'cest_prod_xml', 'orig_icms_xml', 'cst_icms_xml']].rename(columns={'cod_prod_xml': 'CDPRFO'})
df_completa = df_agrelli.merge(df_xml_comp, on='CDPRFO')
df_analise = df_agrelli.reset_index(drop=True) #df.reset_index(drop=True)

### CRIANDO DF VERSAO FINAL  ###
df_vfinal = pd.DataFrame()

#### VER OS INFORMANÇÕES DFS  #####
print(df_agrelli[:0])
print(df_agrelli.info())
print(df_xml_comp[:0])
print(df_xml_comp.info())


##### Função versão 1 COMPARANDO INFORMAÇÕES   ######
def check_item(num_cod):
    ck_var = df_xml_comp[df_xml_comp['CDPRFO'] == str(num_cod)]

    if not ck_var.empty:
        ck_cod_nfe = ck_var._values[0][0]
        ck_nome_nf = ck_var._values[0][1]
        ck_ncm_nf = ck_var._values[0][2]
        ck_cest_nf = ck_var._values[0][3]
        ck_orig_nf = ck_var._values[0][4]
        ck_cst_nf = ck_var._values[0][5]
        print(f'{ck_cod_nfe} | {ck_nome_nf} | {ck_ncm_nf} | {ck_cest_nf} | {ck_orig_nf} | {ck_cst_nf}')
        return f'{ck_cod_nfe} | {ck_nome_nf} | {ck_ncm_nf} | {ck_cest_nf} | {ck_orig_nf} | {ck_cst_nf}'
    else:
        vazio = 'vazio'
        print(vazio)
        return vazio


##### Aplicando função de validar/verificar informações  ######
#df_xml_sped['XML_info'] = df_xml_sped[9].apply(check_item2)
#print(df_xml_sped['XML_info'][:100])

df_vfinal = df_agrelli
df_vfinal['XML_info'] = df_vfinal['CDPRFO'].apply(check_item)
print(df_vfinal['XML_info'][:100])

#### Separando Colunas de Retorno   ######
df_vfinal[['COD_XML','NOME_XML','NCM_XML','CEST_XML','ORIG_XML','CST_XML']] = df_vfinal['XML_info'].str.rsplit(' | ', expand=True, n=6)

print(df_vfinal[:0])
print(df_vfinal[:10])
print(df_vfinal.info())



#### Salvando em Resultado Final em Excel  #####

with pd.ExcelWriter(r'C:\Users\CLAY\PycharmProjects\Proj_inicial\ARQ\ANALISE_XML_CADASTRO_AGRELLI_12-2023.xlsx', mode='w') as writer:
    df_vfinal.to_excel(writer, sheet_name='ANALISE_XML_CADASTRO')