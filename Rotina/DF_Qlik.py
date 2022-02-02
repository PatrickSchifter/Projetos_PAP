import pandas as pd
from Projetos.Rotina.config import path_dir_tod, file, dest_path_qlik, ano_atual, meses_str, mes_atual, file_a, all_index, prov_index, \
    index_qlik, conversor_dt, fatiamento, to_date_time

path_final = 'C:/Users/patrick.paula/Porto a Porto Comercio de IMP e EXP LTDA/Afonso Marcon - QlikSense_Logistica/Teste.xlsx'
path_final_c = 'C:/Users/patrick.paula/Porto a Porto Comercio de IMP e EXP LTDA/Afonso Marcon - QlikSense_Logistica/PlanilhaDataTransporte.xlsx'
path_txt = 'C:/Users/patrick.paula/Porto a Porto Comercio de IMP e EXP LTDA/Afonso Marcon - QlikSense_Logistica/Teste.txt'

df_qlik = pd.read_excel(dest_path_qlik)

df_qlik['N.F'] = df_qlik.rename(columns={'N.F': 'NF'}, inplace=True)
df_mon_m = pd.read_excel(io=file, sheet_name=meses_str[str(int(mes_atual))] + '-' + ano_atual)
try:
    df_mon_m_ant = pd.read_excel(io=file, sheet_name=meses_str[str(int(mes_atual) - 1)] + '-' + ano_atual)
except KeyError:
    df_mon_m_ant = pd.read_excel(io=file_a, sheet_name=meses_str['12'] + '-' + str(int(ano_atual) - 1))
    df_mon_m_ant.columns = all_index

df_mon = pd.concat(objs=[df_mon_m, df_mon_m_ant], ignore_index=True)
df_mon['Número'] = df_mon.rename(columns={'Número': 'NF'}, inplace=True)
df_qlik.drop(columns='N.F', inplace=True)
df_qlik.columns = prov_index
df_qlik = pd.merge(left=df_qlik, right=df_mon, how='left', on='NF')
df_qlik = df_qlik.fillna('-')



# Conversões
df_qlik['DataColeta'] = df_qlik['DataColeta'].astype(dtype='str')
df_qlik['DataAgenda'] = df_qlik['DataAgenda'].astype(dtype='str')
df_qlik['DataEntrega'] = df_qlik['DataEntrega'].astype(dtype='str')
df_qlik['Data-De-Coleta'] = df_qlik['Data-De-Coleta'].astype(dtype='str')
df_qlik['D_Entrega'] = df_qlik['D_Entrega'].astype(dtype='str')
df_qlik['Agendamento'] = df_qlik['Agendamento'].astype(dtype='str')



df_qlik['DataColeta'] = df_qlik['DataColeta'].apply(func=lambda val: conversor_dt(val))
df_qlik['DataAgenda'] = df_qlik['DataAgenda'].apply(func=lambda val: conversor_dt(val))
df_qlik['DataEntrega'] = df_qlik['DataEntrega'].apply(func=lambda val: conversor_dt(val))
df_qlik['Data-De-Coleta'] = df_qlik['Data-De-Coleta'].apply(func=lambda val: conversor_dt(val))
df_qlik['D_Entrega'] = df_qlik['D_Entrega'].apply(func=lambda val: conversor_dt(val))
df_qlik['Agendamento'] = df_qlik['Agendamento'].apply(func=lambda val: conversor_dt(val))

df_qlik['DataColeta'] = df_qlik['DataColeta'] + '!' + df_qlik['Data-De-Coleta']
df_qlik['DataAgenda'] = df_qlik['DataAgenda'] + '!' + df_qlik['Agendamento']
df_qlik['DataEntrega'] = df_qlik['DataEntrega'] + '!' + df_qlik['D_Entrega']

df_qlik['DataColeta'] = df_qlik['DataColeta'].apply(func=lambda val: fatiamento(val))
df_qlik['DataAgenda'] = df_qlik['DataAgenda'].apply(func=lambda val: fatiamento(val))
df_qlik['DataEntrega'] = df_qlik['DataEntrega'].apply(func=lambda val: fatiamento(val))

df_qlik['DataColeta'] = df_qlik['DataColeta'].apply(func=lambda val: to_date_time(val))
df_qlik['DataAgenda'] = df_qlik['DataAgenda'].apply(func=lambda val: to_date_time(val))
df_qlik['DataEntrega'] = df_qlik['DataEntrega'].apply(func=lambda val: to_date_time(val))

df_qlik = df_qlik.fillna('-')

df_qlik = df_qlik[prov_index]
df_qlik.columns = index_qlik

writer = pd.ExcelWriter(path=path_final_c, engine='openpyxl')
df_qlik.to_excel(excel_writer=writer, sheet_name='Sheet1', index=False)
writer.save()

