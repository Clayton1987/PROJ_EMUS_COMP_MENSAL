import pandas as pd
import xmltodict
from tkinter import filedialog , messagebox
from datetime import datetime

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

        # dic_nf = documento['nfeProc']['NFe']['infNFe']
        # if type(documento['nfeProc']['NFe']['infNFe']['det']) == dict:
        #     dic_nf_prod = []
        #     print('++++++++++++++++++++++++++++++++aqui+++++++++++++++++++++++++++')
        #     print(nota)
        #     dic_nf_prod.append(documento['nfeProc']['NFe']['infNFe']['det'])
        # else:
        #     dic_nf_prod = documento['nfeProc']['NFe']['infNFe']['det']
        # #dic_nf_prod = documento['nfeProc']['NFe']['infNFe']['det']

        #### Informações Fornecedor e NF

        inf_notafiscal = documento['nfeProc']['NFe']['infNFe']
        chave_nfe = documento['nfeProc']['NFe']['infNFe']['@Id'].replace('NFe', '')
        info_capa = documento['nfeProc']['NFe']['infNFe']['total']['ICMSTot']

        #### Informações Fornecedor e NF

        dat_nf = inf_notafiscal['ide']['dhEmi'][:10]
        dat_nf = datetime.strptime(dat_nf, '%Y-%m-%d')
        dat_nf = dat_nf.strftime("%d/%m/%Y")
        numero_nf = inf_notafiscal['ide']['nNF']
        try:
            cnpj_emit = inf_notafiscal['emit']['CNPJ'] ##AJUSTES
        except:
            cnpj_emit = inf_notafiscal['emit']['CPF'] ##AJUSTES
        nome_emit = inf_notafiscal['emit']['xNome']

        #### Informações CAPA: Impostos, Totais ###AJUSTES
        bc_icms_nf = info_capa['vBC'].replace('.',',')
        icms_nf = info_capa['vICMS'].replace('.',',')
        ipi_nf = info_capa['vIPI'].replace('.',',')
        valor_nf = info_capa['vNF'].replace('.',',')

        #### Guardando as Informações Dicionario / lista
        notas_xml = []
        lst_resp_notas = []

        ### Dicionario   ###AJUSTES
        dict_resp_nf = {
            'dat_nf': dat_nf,
            'chave_nfe' : chave_nfe,
            'cnpj_emit': cnpj_emit,
            'nome_emit': nome_emit,
            'numero_nf': numero_nf,
            'valor_nf': valor_nf,
            'bc_icms_nf': bc_icms_nf,
            'icms_nf' : icms_nf,
            'ipi_nf': ipi_nf,

        }

        ##### Lista
        lista_nf = [dat_nf, chave_nfe, cnpj_emit, nome_emit, numero_nf, valor_nf, bc_icms_nf, icms_nf, ipi_nf]
        notas_xml.append(lista_nf)
        lst_resp_notas.append(dict_resp_nf)


        # #### Listando Produtos
        # lst_resposta = []
        # lst_resposta2 = []
        # notas_xml = []
        # lst_qdt_itens = '0'
        # produtos = dic_nf_prod
        #
        #
        # ## Criando um lista:
        # dados = [dat_nf,numero_nf,emit_nf,cnpj_nf,cod_prod_xml,nome_prod_xml,ncm_prod_xml,
        #          cest_prod_xml,orig_icms_xml,cst_icms_xml,cst_pis_xml,p_pis_xml,cst_cofins_xml,
        #          p_cofins_xml,vlr_prod_xml,ret_icms_xml,bcmono_icms_xml,adrem_icms_xml]
        #
        # notas_xml.append(dados)
        #
        #
        # lst_resposta2.append(resposta)
        # lst_resposta.append([resposta])
        # # resposta_df = pd.DataFrame(lst_resposta2)
        # # resposta_df = resposta_df._append(resposta_df)
        # # print(resposta_df)
        # # print(lst_resposta2)
        # #dic_resp = {f'item{i}' : resposta}
    elif 'procEventoNFe' in documento.keys(): ##AJUSTES
        print('tipo evento')
        return
    else:
        #pass
        print(documento.keys())
        tipo_nfe = list(documento.keys())
        messagebox.showerror(title="Arquivo inválido, Verifique se é NFe", message=f"Impossível ler Arquivo:\n\n{nota}\n\nTipo do arq: {tipo_nfe[0]}")
    print('lista notas_xml')
    print(*notas_xml, sep="\n")
    print('lista lst_resp_notas')
    return lst_resp_notas


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

# resposta_dfinal = pd.to_datetime(resposta_dfinal['dat_nf'])  #, format='%d/%m/%Y')
# resposta_dfinal = pd.to_datetime(resposta_dfinal['dat_nf'], format='%d-%m-%Y')
print(resposta_dfinal)
print(resposta_dfinal.info())
resposta_dfinal.to_excel(f'{pasta_arqs}/NFs_xml_capa_FILIAL_2021_01-2024.xlsx', sheet_name='XMLs')





