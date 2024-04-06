import pandas as pd
import os #pathlib
from tkinter import filedialog , messagebox


#### LER OS ARQUIVOS SPEDs EXTRAIR O BLOCO ESPECIFICADO PARA EXCEL


df_sped = pd.DataFrame()
df_sped2 = pd.DataFrame()
#df_NFP = pd.read_csv(r'ARQ\ConsultaNFP2.csv', sep=';', encoding='utf-8')
lst_sped = []
lst_df = []
tip_block = 'C100'
mes = '02_2024'


def ler_arq_sped(arq_txt):
    ### Ler os arquivos SPEDs
    df_loop = pd.DataFrame()
    # with open('sped1.txt', 'r',encoding='ANSI') as arquivo:  #US-ASCII #utf-16 #ISO-8895-1
    with open(arq_txt, 'r',encoding='ANSI') as arquivo:  #US-ASCII #utf-16 #ISO-8895-1 #ANSI
        print(arquivo)
        arq_sped = arquivo
        msg = arquivo.readlines()
        num = 0

        var = []
        for linha in msg:
            if linha.startswith(f'|{tip_block}|'): #if '|C100|' in linha:
                num += 1
                lst_sped.append(list([linha]))
                #lst_df.extend(lst_sped)

                # print('lst_sped')
                # print(lst_sped)
                #lst_sped.append([lst_sped])
                #df_sped = df_sped._append([linha], ignore_index=False)
                #df_sped._append(lst_sped)


                # print(df_loop[:10])
                # df_sped._append({'A': [linha]}, ignore_index=True)
                #df_sped = pd.concat([df_sped, df_loop], axis=0, ignore_index=True)
                # df_loop = pd.DataFrame.from_dict([linha])
                # df_sped = df_sped._append(df_loop)
                # resposta_df = pd.DataFrame.from_dict(lst_resp)
                # resposta_dfinal = resposta_dfinal._append(resposta_df)
            # df_sped._append(var)
        # print(len(df_sped))


    return lst_sped
                # print(lst_df)
                # print('df_sped._append([linha])')
                # print(df_sped._append([linha]))
                # print('df_sped2')
                # print(df_sped2)
                # print(linha.count('|'))


#
# print(len(msg))
# print(num)
# print(lst_sped)

##### Pasta Onde arquivos serão lidos

pasta_arqs = filedialog.askdirectory()
print(pasta_arqs)
i=0
lista_arquivos = os.listdir(pasta_arqs)
for item in lista_arquivos:
    if item.endswith('.txt'):
        print(item)
        i+=1
        print(i)
        print(f'{pasta_arqs}/{item}')
        print(ler_arq_sped(f'{pasta_arqs}/{item}'))
        print(len(lst_sped))
        #df_sped2 = df_sped2.append()


print('lst_sped')
print(lst_sped)
print('lst_df')
print(lst_df)
print(len(lst_sped))
df_sped = df_sped._append(lst_sped)

# print('lista_arquivos')
# print(len(lista_arquivos))
#
# print('lst_sped')
# print(lst_sped)
print('df_sped')
print(df_sped[:10])
print(df_sped2[:10])

# ### Finalizando DF
# print(df_sped[:10])
df_sped = df_sped[0].str.split('|', expand=True)
# #df_sped = df['Local'].str.split('|', expand=True)
print(df_sped[:10])
#print(df_NFP[:10])
# print(df_sped.info())
# ### Salvando Excel
#df_sped.to_excel('ARQ_SPED.xlsx', sheet_name='SPEDS')

with pd.ExcelWriter(rf'{pasta_arqs}\ARQ_{mes}_SPED-{tip_block}.xlsx', mode='w') as writer:
    #df1.to_excel(writer, sheet_name='Sheet_name_3')
    df_sped.to_excel(writer, sheet_name=f'{tip_block}')
    #df_NFP.to_excel(writer, sheet_name='NFP', encoding='ANSI')


###############################################################


# ##### Pasta Onde arquivo serão lidos
# import os   #pathlib
#
# pasta_arqs = filedialog.askdirectory()
# print(pasta_arqs)
# lista_arquivos = os.listdir(pasta_arqs)
#
# resposta_dfinal = pd.DataFrame()
# for arquivo in lista_arquivos:
#     if '.xml' in arquivo:
#         print('passo arquivo')
#         print(arquivo)
#         #lst_resposta2 = ler_xml_nota(f'arquivos/NFs Finais/{arquivo}')
#         #print(ler_xml_nota(f'arquivos/NFs Finais/{arquivo}'))
#         lst_resp = ler_xml_nota(f'{pasta_arqs}/{arquivo}')
#         resposta_df = pd.DataFrame.from_dict(lst_resp)
#         resposta_dfinal = resposta_dfinal._append(resposta_df)
#     else:
#         pass
#
#
#     #print(resposta_dfinal)
#
#
# print(resposta_dfinal)
# resposta_dfinal.to_excel(f'{pasta_arqs}/NFs_xml.xlsx', sheet_name='XMLs')