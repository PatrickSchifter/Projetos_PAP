import pandas as pd
import time
st_time = time.time()
t_prazo_entrega = r'C:/Users/patrick.paula/Porto a Porto Comercio de IMP e EXP LTDA/Afonso Marcon - QlikSense_Logistica/PrazoDeEntrega.xlsx'
df_lead = pd.read_excel(io=t_prazo_entrega)
df_lead.columns = ['Cd_Cidade', 'Cidade-Destinatário', 'Uf', 'Cd_Transportador', 'Fantasia_Do_Transportador', 'Lead_Time']
df_lead['Cidade-Destinatário'] = df_lead['Cidade-Destinatário'].astype(dtype='str')
df_lead['Cidade-Destinatário'] = df_lead['Cidade-Destinatário'].apply(func=lambda val: val.upper())
# print(df_lead)
# print(time.time() - st_time)

