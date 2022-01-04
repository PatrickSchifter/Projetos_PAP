import pandas as pd
import os
from Variaveis import dir_coletas, dir_arq_coletas
import shutil

dfs_cols = []
for file in list(os.listdir(dir_coletas)):
    if '.xlsx' in file or '.xls' in file:
        arq = dir_coletas + file
        df_prov = pd.read_excel(io=arq)
        df_prov.columns = [['Número', 'Data-De-Coleta']]
        dfs_cols.append(df_prov)
        shutil.move(src=arq, dst=dir_arq_coletas + file)


df_col = pd.concat(objs=dfs_cols)

