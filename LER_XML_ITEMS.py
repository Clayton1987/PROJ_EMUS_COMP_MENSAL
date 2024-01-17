import pandas as pd
import xmltodict
from tkinter import filedialog , messagebox

####   Importar produtos excel   ###
# lst_prod_sist = pd.read_excel(r'arquivos/AGRELI_ANALISE.xlsx', sheet_name='EFAPRODU')
# lst_prod_sist = lst_prod_sist[['CODPRO','NOMPRO','CLAFIS','CEST','ALQICM','CODTRI','CODTR2','REDAGR','DTCUST','SALEST','SALDEP','REFERE','CDPRFO']]
# print(lst_prod_sist[:10])

def ler_xml_nota(nota):
    #####  Importando XML Produtos/fornecedor
    #with open(r'arquivos/NFs Finais/35231187235172000807550010009150261251855739.xml', 'rb') as arquivo:
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    if 'nfeProc' in documento:
        dic_nf = documento['nfeProc']['NFe']['infNFe']
        if type(documento['nfeProc']['NFe']['infNFe']['det']) == dict:
            dic_nf_prod = []
            dic_nf_prod.append(documento['nfeProc']['NFe']['infNFe']['det'])
        else:
            dic_nf_prod = documento['nfeProc']['NFe']['infNFe']['det']
        #dic_nf_prod = documento['nfeProc']['NFe']['infNFe']['det']

        #### Informações Fornecedor e NF
        dat_nf = dic_nf['ide']['dhEmi']
        numero_nf = dic_nf['ide']['nNF']
        emit_nf = dic_nf['emit']['xNome']
        #cnpj_nf = dic_nf['emit']['CNPJ']
        try:
            cnpj_nf = dic_nf['emit']['CNPJ'] ##AJUSTES
        except:
            cnpj_nf = dic_nf['emit']['CPF'] ##AJUSTES


        #### Listando Produtos
        lst_resposta = []
        lst_resposta2 = []
        notas_xml = []
        lst_qdt_itens = '0'
        produtos = dic_nf_prod
        for i, produto in enumerate(produtos):
            print(produto)
            cod_prod_xml = produto['prod']['cProd']
            nome_prod_xml = produto['prod']['xProd']
            ncm_prod_xml = produto['prod']['NCM']
            vlr_prod_xml = produto['prod']['vProd'].replace('.',',')
            try:
                cest_prod_xml = produto['prod']['CEST']
            except:
                cest_prod_xml = 'NN'
            print(produto['imposto']['ICMS'].keys())
            if 'ICMS00' in produto['imposto']['ICMS']:
                icms = 'ICMS00'
            else:
                lst_icms = list(produto['imposto']['ICMS'].keys())
                #print(lst_icms)
                icms = lst_icms[0]
                print(icms)
            orig_icms_xml = produto['imposto']['ICMS'][f'{icms}']['orig']
            try:
                cst_icms_xml = produto['imposto']['ICMS'][f'{icms}']['CST']
            except:
                lst_i_cst = list(produto['imposto']['ICMS'][f'{icms}'].keys())
                #print(lst_icms)
                i_cst = lst_i_cst[1]
                print('cst')
                print(i_cst)
                cst_icms_xml = produto['imposto']['ICMS'][f'{icms}'][f'{i_cst}']
            #p_icms_xml = produto['imposto']['ICMS'][f'{icms}']['pICMS']
            if 'PISAliq' in produto['imposto']['PIS']:
                pis_al = 'PISAliq'
            else:
                lst_pis_al = list(produto['imposto']['PIS'].keys())
                pis_al = lst_pis_al[0]
                print('pis_al')
                print(pis_al)
            cst_pis_xml = produto['imposto']['PIS'][f'{pis_al}']['CST']
            try:
                p_pis_xml = produto['imposto']['PIS'][f'{pis_al}']['pPIS']
            except:
                lst_p_pis = list(produto['imposto']['PIS'][f'{pis_al}'].keys())
                try:
                    p_pis = lst_p_pis[3]
                except:
                    p_pis = 'NN'
                print('pPIS')
                print(p_pis)
                try:
                    p_pis_xml = produto['imposto']['PIS'][f'{pis_al}'][f'{p_pis}']
                except:
                    p_pis_xml = 'NN'

            if 'COFINSAliq' in produto['imposto']['COFINS']:
                cofins_al = 'COFINSAliq'
            else:
                lst_cofis_al = list(produto['imposto']['COFINS'].keys())
                cofins_al = lst_cofis_al[0]
                print('cofins_Al')
                print(cofins_al)

            cst_cofins_xml = produto['imposto']['COFINS'][f'{cofins_al}']['CST']
            try:
                p_cofins_xml = produto['imposto']['COFINS'][f'{cofins_al}']['pCOFINS']
            except:
                lst_pcofins = list(produto['imposto']['COFINS'][f'{cofins_al}'].keys())
                try:
                    p_cofins = lst_pcofins[3]
                except:
                    p_cofins = 'NN'
                print('pcofins')
                print(p_cofins)
                try:
                    p_cofins_xml = produto['imposto']['COFINS'][f'{cofins_al}'][f'{p_cofins}']
                except:
                    p_cofins_xml = 'NN'
            try:
                ret_icms_xml = produto['imposto']['ICMS'][f'{icms}']['vICMSMonoRet'].replace('.',',')
                bcmono_icms_xml = produto['imposto']['ICMS'][f'{icms}']['qBCMonoRet'].replace('.',',')
                adrem_icms_xml = produto['imposto']['ICMS'][f'{icms}']['adRemICMSRet'].replace('.', ',')
            except:
                ret_icms_xml = 'NN'
                bcmono_icms_xml = 'NN'
                adrem_icms_xml = 'NN'

            #print(i)


            resposta = {
                'dat_nf' : dat_nf,
                'numero_nf' : numero_nf,
                'emit_nf' : emit_nf,
                'cnpj_nf' : cnpj_nf,
                'cod_prod_xml' : cod_prod_xml,
                'nome_prod_xml' : nome_prod_xml,
                'ncm_prod_xml' : ncm_prod_xml,
                'cest_prod_xml' : cest_prod_xml,
                'orig_icms_xml' : orig_icms_xml,
                'cst_icms_xml' : cst_icms_xml,
                #'p_icms_xml' : p_icms_xml,
                'cst_pis_xml' : cst_pis_xml,
                'p_pis_xml' : p_pis_xml,
                'cst_cofins_xml' : cst_cofins_xml,
                'p_cofins_xml' : p_cofins_xml,
                'vlr_prod_xml' : vlr_prod_xml,
                'ret_icms_xml' : ret_icms_xml,
                'bcmono_icms_xml' : bcmono_icms_xml,
                'adrem_icms_xml' : adrem_icms_xml,


            }
            resposta2 = {
                'dat_nf' : dat_nf,
                'numero_nf' : numero_nf,
                'emit_nf' : emit_nf,
                'cnpj_nf' : cnpj_nf,
                'cod_prod_xml' : cod_prod_xml,
                'nome_prod_xml' : nome_prod_xml,
                'ncm_prod_xml' : ncm_prod_xml,
                'cest_prod_xml': cest_prod_xml,
                'orig_icms_xml' : orig_icms_xml,
                'cst_icms_xml' : cst_icms_xml,
                #'p_icms_xml' : p_icms_xml,
                'cst_pis_xml' : cst_pis_xml,
                'p_pis_xml' : p_pis_xml,
                'cst_cofins_xml' : cst_cofins_xml,
                'p_cofins_xml' : p_cofins_xml,
                'vlr_prod_xml': vlr_prod_xml,
                'ret_icms_xml': ret_icms_xml,
                'bcmono_icms_xml': bcmono_icms_xml,
                'adrem_icms_xml': adrem_icms_xml,

            }
            # print(resposta)

            ## Criando um lista:
            dados = [dat_nf,numero_nf,emit_nf,cnpj_nf,cod_prod_xml,nome_prod_xml,ncm_prod_xml,
                     cest_prod_xml,orig_icms_xml,cst_icms_xml,cst_pis_xml,p_pis_xml,cst_cofins_xml,
                     p_cofins_xml,vlr_prod_xml,ret_icms_xml,bcmono_icms_xml,adrem_icms_xml]

            notas_xml.append(dados)


            lst_resposta2.append(resposta)
            lst_resposta.append([resposta])
            # resposta_df = pd.DataFrame(lst_resposta2)
            # resposta_df = resposta_df._append(resposta_df)
            # print(resposta_df)
            # print(lst_resposta2)
            #dic_resp = {f'item{i}' : resposta}
    elif 'procEventoNFe' in documento.keys(): ##AJUSTES
        print('tipo evento')
        return
    else:
        #pass
        print(documento.keys())
        tipo_nfe = list(documento.keys())
        messagebox.showerror(title="Arquivo inválido, Verifique se é NFe", message=f"Impossível ler Arquivo:\n\n{nota}\n\nTipo do arq: {tipo_nfe[0]}")
    print(*notas_xml, sep="\n")
    return lst_resposta2


##### Pasta Onde arquivo serão lidos
import os   #pathlib

pasta_arqs = filedialog.askdirectory()
print(pasta_arqs)
lista_arquivos = os.listdir(pasta_arqs)

resposta_dfinal = pd.DataFrame()
for arquivo in lista_arquivos:
    if '.xml' in arquivo:
        print('passo arquivo')
        print(arquivo)
        #lst_resposta2 = ler_xml_nota(f'arquivos/NFs Finais/{arquivo}')
        #print(ler_xml_nota(f'arquivos/NFs Finais/{arquivo}'))
        lst_resp = ler_xml_nota(f'{pasta_arqs}/{arquivo}')
        resposta_df = pd.DataFrame.from_dict(lst_resp)
        resposta_dfinal = resposta_dfinal._append(resposta_df)
    else:
        pass


    #print(resposta_dfinal)


print(resposta_dfinal)
resposta_dfinal.to_excel(f'{pasta_arqs}/NFs_xml_ITEM_19A23.xlsx', sheet_name='XML_ITEM')





