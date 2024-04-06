import pandas as pd
from tkinter import filedialog
import os

class funcs:
    def unir_arq():

        df_unido_xml = pd.DataFrame()
        pasta_arqs2 = filedialog.askopenfiles()
        #print(pasta_arqs2[0].name)
        caminho_dir = os.path.dirname(pasta_arqs2[0].name)
        print(caminho_dir)


        for arquivo in pasta_arqs2:
            print(arquivo.name)

            if arquivo.name.endswith('xlsx'):
                excel_arq = pd.read_excel(arquivo.name)
                #print(excel_arq)
                df_unido_xml = df_unido_xml._append(excel_arq)
                print(df_unido_xml.head(20))
        #df_unido_xml.to_excel(f'{caminho_dir}/NFs_xml_capa_FILIAL_2021_02-2024_unido.xlsx', sheet_name='XMLs')
        df_unido_xml.to_excel(f'{caminho_dir}/CTEs_xml_capa_FILIAL_2021_02-2024_unido.xlsx', sheet_name='CTE_XMLs', index=False)
        #return print(df_unido_xml.info())

funcs.unir_arq()