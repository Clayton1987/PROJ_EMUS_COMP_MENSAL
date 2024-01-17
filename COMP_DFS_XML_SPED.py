#### Lendo Excel comparando informações
### data criação: 10/12/2023
### modificado: 10/12/2023
import pandas as pd
import time


###C:\Users\CLAY\Documents\AUTO_MAG\ANALISE\ARQ
###C:\Users\CLAY\Documents\EMUSA\MATRIZ_XML_NF_2019_2023\MATRIZ_XML_NF_2019_2023

#### Importando arquivos
# df_sped = pd.read_excel(r'ARQ\ARQ_SPED_C100_MATRIZ.xlsx', sheet_name='SPEDS')
# df_xmls = pd.read_excel(r'ARQ\NFs_xml_capa.xlsx', sheet_name='XMLs')
df_sped = pd.read_excel(r'C:\Users\CLAY\Documents\EMUSA\SPED_MATRIZ_2019_2023\12-2023\ARQ_MATRIZ_MES_12_SPED-C100.xlsx', sheet_name='C100')
df_xmls = pd.read_excel(r'C:\Users\CLAY\Documents\EMUSA\MATRIZ_XML_NF_2019_2023\NFs_xml_capa_FILIAL_2021_01-2024_160124.xlsx', sheet_name='XMLs')

######    Limpar colunas Importantes VAZIASSS    #######
df_sped.dropna(subset=[9, 10, 11])

######   Tratando tipos das COLUNAS    ###########
#df_sped[9] = df_sped[9].astype(int)
df_sped[8] = df_sped[8].astype(int)
#df_sped[9] = pd.to_numeric(df_sped[9], downcast="integer")
#df_sped[8] = df_sped[8].fillna(0).astype(int)
df_sped[10] = pd.to_numeric(df_sped[10], downcast="integer")
df_sped[11] = pd.to_numeric(df_sped[11], downcast="integer")
df_sped[12] = df_sped[12].str.replace(',','.')
df_sped[12] = pd.to_numeric(df_sped[12], downcast="float")
df_sped['CNPJ'] = df_sped[9].str[6:20]._values
df_sped['NUM_NF'] = df_sped[9].str[25:34]._values
df_sped = df_sped[df_sped[3] == 1]

print(df_xmls.info())


#### Criando df Final da Informações #####
df_xml_sped = pd.DataFrame()
df_xml_sped = df_xml_sped._append(df_sped, ignore_index=True)


##### Função versão 1 COMPARANDO INFORMAÇÕES   ######
def check_item(num_chave):
    ck_var = df_xmls[df_xmls['chave_nfe'] == str(num_chave)]

    if not ck_var.empty:
        ck_chave_nfe = ck_var._values[0][2]
        ck_dat_nf = ck_var._values[0][1]
        ck_valor_nf = ck_var._values[0][6]
        ck_bc_icms_nf = ck_var._values[0][7]
        ck_icms_nf = ck_var._values[0][8]
        ck_ipi_nf = ck_var._values[0][9]
        print(f'{ck_chave_nfe} | {ck_dat_nf} | {ck_valor_nf} | {ck_bc_icms_nf} | {ck_icms_nf} | {ck_ipi_nf}')
        return f'{ck_chave_nfe} | {ck_dat_nf} | {ck_valor_nf} | {ck_bc_icms_nf} | {ck_icms_nf} | {ck_ipi_nf}'
    else:
        vazio = 'vazio'
        print(vazio)
        return vazio


##### Função versão  2   ######
def check_item2(num_chave):
    ck_var = df_xmls[df_xmls['chave_nfe'] == str(num_chave)]

    if not ck_var.empty:
        ck_chave_nfe = ck_var._values[0][2]
        ck_dat_nf = ck_var._values[0][1]
        ck_valor_nf = ck_var._values[0][6]
        ck_bc_icms_nf = ck_var._values[0][7]
        ck_icms_nf = ck_var._values[0][8]
        ck_ipi_nf = ck_var._values[0][9]
        print(f'{ck_chave_nfe} | {ck_dat_nf} | {ck_valor_nf} | {ck_bc_icms_nf} | {ck_icms_nf} | {ck_ipi_nf}')
        return f'{ck_chave_nfe} | {ck_dat_nf} | {ck_valor_nf} | {ck_bc_icms_nf} | {ck_icms_nf} | {ck_ipi_nf}'
    else:
        vazio = 'vazio'
        print(vazio)
        return vazio


##### Aplicando função de validar/verificar informações  ######
df_xml_sped['XML_info'] = df_xml_sped[9].apply(check_item)
print(df_xml_sped['XML_info'][:100])

#### Separando Colunas de Retorno   ######
df_xml_sped[['CHAVE_NF', 'DAT_NF', 'VALOR_NF', 'BC_ICMS_NF', 'ICMS_NF', 'IPI_NF']] = df_xml_sped['XML_info'].str.rsplit(' | ',expand=True, n=6)

print(df_xml_sped[:0])
print(df_xml_sped[:10])
print(df_xml_sped.info())


#### Salvando em Resultado Final em Excel  #####

with pd.ExcelWriter(r'C:\Users\CLAY\Documents\EMUSA\ANALISE_SPED_XML_ICMS_160124.xlsx', mode='w') as writer:
    df_xml_sped.to_excel(writer, sheet_name='ANALISE_ICMS')



################################################
# num_chave = 13190521338912000148550040000003561005785550
# print(df_xml_sped[df_xml_sped[9] == str(num_chave)])
#
# print(df_xmls[df_xmls['chave_nfe'] == str(num_chave)])
# print(df_xmls['chave_nfe'].isin([num_chave]))
# print(df_xmls[df_xmls['chave_nfe'] == str(num_chave)]._values)
# #print(df_xmls[df_xmls['chave_nfe'] == str(num_chave)]._values[0][1])
#
# for item in range(100):
#     print(df_xmls['chave_nfe'][item])
#     print('+++++++++++++++++++++')
#     check_item(df_xmls['chave_nfe'][item])










