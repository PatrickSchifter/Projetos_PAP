import pandas as pd
from pandas import Series
import os
from config import dir_coletas, dir_arq_coletas, today
import shutil


def conversor(val):
    val = val.split('-')
    valor = val[2] + '/' + val[1] + '/' + val[0]
    return valor

dfs_cols = []
for file in list(os.listdir(dir_coletas)):
    if '.xlsx' in file or '.xls' in file:
        arq = dir_coletas + file
        df_prov = pd.read_excel(io=arq)
        df_prov.columns = ['Número', 'Data-De-Coleta_col']
        dfs_cols.append(df_prov)
        # shutil.move(src=arq, dst=dir_arq_coletas + today + file)
try:
    df_col = pd.concat(objs=dfs_cols, ignore_index=True, keys=['Número', 'Data-De-Coleta_col'])
    df_col['Data-De-Coleta_col'] = df_col['Data-De-Coleta_col'].astype('str')
    df_col['Data-De-Coleta_col'] = df_col['Data-De-Coleta_col'].apply(func=lambda val: conversor(val))

except ValueError:
    pass
