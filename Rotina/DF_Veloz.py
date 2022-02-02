import pandas as pd
import os
from Projetos.Rotina.config import dest_path, meses, mes_atual, column_veloz, conversor_dt, dia_hoje

"""Até o dia 10 deve-se fazer a conferência do mes atual e mes - 1
    Verificar como será o padrão da Veloz ao virar o ano"""

path_plan = dest_path + '/'
for itens in list(os.listdir(dest_path)):
    if 'CONTROLE DE PEDIDOS PAP' in itens:
        path_plan = path_plan + itens
        break

df_veloz = pd.ExcelFile(path_plan)
if dia_hoje == 1:
    if int(mes_atual) < 10:
        ind_mes = '0' + str(int(mes_atual) - 1)
    else:
        ind_mes = mes_atual
else:
    ind_mes = mes_atual

sheets = df_veloz.sheet_names

for x in sheets:
    if x == meses[ind_mes].title():
        df_veloz = pd.read_excel(io=path_plan, sheet_name=x)
    elif x == meses[ind_mes].upper():
        df_veloz = pd.read_excel(path_plan, sheet_name=x)
    elif x == meses[ind_mes]:
        df_veloz = pd.read_excel(path_plan, sheet_name=x)


df_veloz.columns = column_veloz
df_veloz = df_veloz[['Número', 'Data_De_Coleta']]
df_veloz = df_veloz.fillna('-')
df_veloz = df_veloz.drop([0, 1, 2])
df_veloz['Data_De_Coleta'] = df_veloz['Data_De_Coleta'].astype(dtype='str')
df_veloz['Data_De_Coleta'] = df_veloz['Data_De_Coleta'].apply(func=lambda val: conversor_dt(val))

if mes_atual != '1':
    try:
        df_veloz_m = pd.read_excel(io=path_plan, sheet_name=meses[str(int(mes_atual) - 1)].title())
        df_veloz_m.columns = column_veloz
        df_veloz_m = df_veloz_m[['Número', 'Data_De_Coleta']]
        df_veloz_m = df_veloz_m.fillna('-')
        df_veloz_m = df_veloz_m.drop([0, 1, 2])
        df_veloz_m['Data_De_Coleta'] = df_veloz_m['Data_De_Coleta'].astype(dtype='str')
        df_veloz_m['Data_De_Coleta'] = df_veloz_m['Data_De_Coleta'].apply(func=lambda val: conversor_dt(val))
    except KeyError:
        #Exceção já tratada no if
        pass
