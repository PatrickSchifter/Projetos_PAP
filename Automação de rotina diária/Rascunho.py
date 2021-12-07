import pandas as pd
dados = {'Nota': ['456123', '456124', '456125', '456128', '456159', '456160'], 'D_Entrega': ['-', '27/11/2021', '-', '-', '28/11/2021', None]}
dados2 = {'Nota': ['456123', '456128', '456125', '456159', '456160'], 'Data': ['24/11/2021', '25/11/2021', '26/11/2021', '28/11/2021', None], 'Status': ['24/11/2021', '25/11/2021', '26/11/2021', '28/11/2021', None]}

df = pd.DataFrame(data=dados)
df2 = pd.DataFrame(data=dados2)

df3 = pd.merge(left=df, right=df2, how='left')

fil_ent = df['D_Entrega'] != '-'
fil_dat = df2['Data'] > '1'
df4 = df[fil_ent]
df5 = df3[(fil_dat) | (fil_ent)]

df5['n_DATE'] = df5['Data'] + df5['D_Entrega']
print(df5)
