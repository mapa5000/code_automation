# code_automation

import pandas as pd
import numpy as np
import openpyxl
import os

path = r"P:\Houston Gas Drafting\GIS\Corrections\Submitted\2021\Marco Portillo\pending EOP Data Exploration\test\BEx_RouteSmart_pt_Facet_Join.xlsx"

dataframe = pd.read_excel(path)
dataframe.columns

dataframe.dtypes


dataframe['House'] = dataframe['AddNum'].astype(str).map(lambda x:x.split('.')[0])

df = dataframe[['House','Side','StName','StType','First_Name','Last_Name','FACET_NAME']]
df.head()

df.sort_values(by=['FACET_NAME','Last_Name']).head(2)


split_values = df['FACET_NAME'].unique()
print(split_values)

for value in split_values:
    df1 = df[df['FACET_NAME']==value]
    output_file_name = 'FACET_'+str(value)+'.xlsx'
    df1.to_excel(output_file_name,index='False')


os.chdir(r"P:\Houston Gas Drafting\GIS\Corrections\Submitted\2021\Marco Portillo\pending EOP Data Exploration\test")
os.getcwd()
os.listdir()[2:-4]


# ya este me da error ...pero creo que tengo que arreglar la data
files = os.listdir()[2:-1]
for file in files:
    wb_obj = openpyxl.load_workbook(file)
    sheet_obj = wb_obj.active
    cell_obj = sheet_obj.cell(row = 2, column = 1)
    print(cell_obj.value)


# este es el codigo que estaba usando para separar un sefet en multiples sheets
df = pd.DataFrame(np.arange(100).reshape((10, 10)))

writer = pd.ExcelWriter(r"P:\Houston Gas Drafting\GIS\Corrections\Submitted\2021\Marco Portillo\pending EOP Data Exploration\test\test_facet_split\FACET_F32100870.xlsx")
for key, grp in df.groupby(df.index // 2):
    grp.to_excel(writer, f'sheet_{key}', header=False)
    writer.save()
