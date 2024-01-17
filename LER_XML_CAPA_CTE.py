import pandas as pd
import xmltodict
from tkinter import filedialog , messagebox
from datetime import datetime
from tqdm import tqdm

####   Importar produtos excel   ###
# lst_prod_sist = pd.read_excel(r'arquivos/AGRELI_ANALISE.xlsx', sheet_name='EFAPRODU')
# lst_prod_sist = lst_prod_sist[['CODPRO','NOMPRO','CLAFIS','CEST','ALQICM','CODTRI','CODTR2','REDAGR','DTCUST','SALEST','SALDEP','REFERE','CDPRFO']]
# print(lst_prod_sist[:10])

def ler_xml_nota(nota):
    #####  Importando XML Produtos/fornecedor
    #with open(r'arquivos/NFs Finais/35231187235172000807550010009150261251855739.xml', 'rb') as arquivo:
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    if 'cteProc' in documento:

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

        #### Informações da CTe Necessarias
        info_cte = documento['cteProc']['CTe']['infCte']
        chave_cte = documento['cteProc']['CTe']['infCte']['@Id'].replace('CTe', '')
        dat_cte = info_cte['ide']['dhEmi'][:10]
        dat_cte = datetime.strptime(dat_cte, '%Y-%m-%d')
        dat_cte = dat_cte.strftime("%d/%m/%Y")
        num_cte = info_cte['ide']['nCT']
        cfop_cte = info_cte['ide']['CFOP']
        natop_cte = info_cte['ide']['natOp']
        emit_cnpj_cte = info_cte['emit']['CNPJ']
        emit_nome_cte = info_cte['emit']['xNome']
        try:
            rem_cnpj_cte = info_cte['rem']['CNPJ']
        except:
            rem_cnpj_cte = info_cte['rem']['CPF']
        remet_nome_cte = info_cte['rem']['xNome']
        total_serv_cte = info_cte['vPrest']['vTPrest'].replace('.',',')

        try:
            if 'infCTeNorm' in info_cte:
                # if 'infOutros' in info_cte['infCTeNorm']['infDoc']:
                #     nfe_chave_cte = 'AEREO'
                try:
                    nfe_chave_cte = info_cte['infCTeNorm']['infDoc']['infNFe']['chave']
                except:
                    nfe_chave_cte = 'OUTROS'
            else:
                try:
                    nfe_chave_cte = info_cte['infCteComp']['chCTe']
                except:
                    nfe_chave_cte = info_cte['infCteComp']['chave']
                #nfe_chave_cte = info_cte['infCTeNorm']['infDoc']['infNFe']['chave']
        except:
            nfe_chave_cte = info_cte['infCTeNorm']['infDoc']['infNFe']#[0]['chave']

        if 'ICMS00' in info_cte['imp']['ICMS']:
            cst_cte = info_cte['imp']['ICMS']['ICMS00']['CST']
            bcicms_cte = info_cte['imp']['ICMS']['ICMS00']['vBC'].replace('.',',')
            icms_aliq_cte = info_cte['imp']['ICMS']['ICMS00']['pICMS'].replace('.',',')
            icms_valor_cte = info_cte['imp']['ICMS']['ICMS00']['vICMS'].replace('.',',')
        elif 'ICMS90' in info_cte['imp']['ICMS']:
            cst_cte = info_cte['imp']['ICMS']['ICMS90']['CST']
            bcicms_cte = info_cte['imp']['ICMS']['ICMS90']['vBC'].replace('.',',')
            icms_aliq_cte = info_cte['imp']['ICMS']['ICMS90']['pICMS'].replace('.',',')
            icms_valor_cte = info_cte['imp']['ICMS']['ICMS90']['vICMS'].replace('.',',')
        elif 'ICMS60' in info_cte['imp']['ICMS']:
            cst_cte = info_cte['imp']['ICMS']['ICMS60']['CST']
            bcicms_cte = info_cte['imp']['ICMS']['ICMS60']['vBCSTRet'].replace('.',',')
            icms_aliq_cte = info_cte['imp']['ICMS']['ICMS60']['pICMSSTRet'].replace('.',',')
            icms_valor_cte = info_cte['imp']['ICMS']['ICMS60']['vICMSSTRet'].replace('.',',')
        elif 'ICMSOutraUF' in info_cte['imp']['ICMS']:
            cst_cte = info_cte['imp']['ICMS']['ICMSOutraUF']['CST']
            bcicms_cte = info_cte['imp']['ICMS']['ICMSOutraUF']['vBCOutraUF'].replace('.',',')
            icms_aliq_cte = info_cte['imp']['ICMS']['ICMSOutraUF']['pICMSOutraUF'].replace('.',',')
            icms_valor_cte = info_cte['imp']['ICMS']['ICMSOutraUF']['vICMSOutraUF'].replace('.',',')
        elif 'ICMS45' in info_cte['imp']['ICMS']:
            cst_cte = info_cte['imp']['ICMS']['ICMS45']['CST']
            bcicms_cte = 0
            icms_aliq_cte = 0
            icms_valor_cte = 0
        else:  # {'ICMSSN': {'CST': '90', 'indSN': '1'}}
            try:
                cst_cte = info_cte['imp']['ICMS']['ICMSSN']['CST']
            except:
                cst_cte = info_cte['imp']['ICMS']['ICMSSN']['indSN']
            bcicms_cte = 0
            icms_aliq_cte = 0
            icms_valor_cte = 0


        # #############  TESTE
        # nfe_chave_cte = 0
        # bcicms_cte = 0
        # icms_aliq_cte = 0
        # icms_valor_cte = 0
        #### Guardando as Informações Dicionario / lista
        ctes_xml = []
        lst_resp_ctes = []

        #### Dicionario
        dict_resp_cte = {
            'num_cte': num_cte,
            'dat_cte': dat_cte,
            'cfop_cte': cfop_cte,
            'cst_cte': cst_cte,
            'total_serv_cte': total_serv_cte,
            'bcicms_cte': bcicms_cte,
            'icms_aliq_cte': icms_aliq_cte,
            'icms_valor_cte': icms_valor_cte,
            'chave_cte': chave_cte,
            'natop_cte': natop_cte,
            'emit_cnpj_cte': emit_cnpj_cte,
            'emit_nome_cte': emit_nome_cte,
            'rem_cnpj_cte': rem_cnpj_cte,
            'remet_nome_cte': remet_nome_cte,
            'nfe_chave_cte': nfe_chave_cte,

        }

        #### Lista
        ctes_lst = [num_cte, dat_cte, cfop_cte, cst_cte, total_serv_cte, bcicms_cte,  icms_aliq_cte, icms_valor_cte, chave_cte,
                    natop_cte, emit_cnpj_cte, emit_nome_cte, rem_cnpj_cte, remet_nome_cte, nfe_chave_cte]
        ctes_xml.append(ctes_lst)
        lst_resp_ctes.append(dict_resp_cte)



    elif 'procEventoNFe' in documento.keys(): ##AJUSTES
        print('tipo evento')
        return
    else:
        #pass
        print(documento.keys())
        tipo_nfe = list(documento.keys())
        messagebox.showerror(title="Arquivo inválido, Verifique se é NFe", message=f"Impossível ler Arquivo:\n\n{nota}\n\nTipo do arq: {tipo_nfe[0]}")
    # print('lista ctes_xml')
    # print(*ctes_xml, sep="\n")
    # print('lista lst_resp_ctes')
    return lst_resp_ctes


##### Pasta Onde arquivo serão lidos
import os   #pathlib


pasta_arqs = filedialog.askdirectory()
print(pasta_arqs)
lista_arquivos = os.listdir(pasta_arqs)
pbar = tqdm(total=len(lista_arquivos), position=0, leave=True)

resposta_dfinal = pd.DataFrame()
for arquivo in lista_arquivos:
    pbar.update()
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
resposta_dfinal.to_excel(f'{pasta_arqs}/CTEs_xml_capa_01-2024.xlsx', sheet_name='CTE_XMLs')





