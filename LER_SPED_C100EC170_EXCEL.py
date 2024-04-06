import pandas as pd
import os #pathlib
from tkinter import filedialog

#arq_name = r'C:\Users\CLAY\Documents\EMUSA\APURACAO_MENSAL\EFD_CONTRIBUICOES012024.txt'
#arq_name = r'C:\Users\CLAY\Documents\EMUSA\SPED_FILIAL_2021_2023\02-2024\EFD022024.txt'

arq_name = filedialog.askopenfile()
caminho = os.path.dirname(arq_name.name)
arq_name = rf'{arq_name.name}'
print(caminho)

with open(arq_name, 'r', encoding='ANSI') as arq_spd:
    arq_spd1 = arq_spd.readlines()
    # for linha in arq_spd.readlines():
    #     print(linha, end='\n')

    lst_c100_c170 = []
    for item in arq_spd1:
        if item.startswith('|C100|'):
            sep_ln = item.split('|')
            # print(f'|{sep_ln[1]}|{sep_ln[8]}')
            # ln_C100 = f'|{sep_ln[1]}|{sep_ln[8]}|{sep_ln[9]}'
            ln_C100 = f'|{sep_ln[1]}|{sep_ln[8]}'
        if item.startswith('|C170|'):
            # sep_ln = item.split('|')
            nw_item = f'{ln_C100}{item}'
            # print(nw_item)
            lst_c100_c170.append(nw_item)

    print(*lst_c100_c170, sep='\n')

df_spd = pd.DataFrame(lst_c100_c170)
print(df_spd)
df_spd2 = df_spd[0].str.split('|', expand=True)
print(df_spd2.head(150))

with pd.ExcelWriter(rf'{caminho}\EFD_CONTRIBUICOES022024_tst2.xlsx', mode='w') as writer:
    #df1.to_excel(writer, sheet_name='Sheet_name_3')
    df_spd2.to_excel(writer, sheet_name='C100_C170')