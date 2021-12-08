import pandas as pd

dados = {'Nota': ['456123', '456124', '456125', '456128', '456159', '456160', '456161', '456162', '456163'],
         'D_Entrega': ['-', '27/11/2021', '-', '-', '28/11/2021', '-', '-', '-', '-']}
dados2 = {'Nota': ['456123', '456128', '456125', '456159'],
          'Data': ['24/11/2021', '25/11/2021', '26/11/2021', '28/11/2021']}

df = pd.DataFrame(data=dados)
df2 = pd.DataFrame(data=dados2)

def change_data(cel):
    try:
        if len(cel) > 11:
            return cel[-11:]
        else:
            cel = '-'
            return '-'
    except:
        pass

df3 = pd.merge(left=df, right=df2, how='left')

df3.fillna(value='-', inplace=True)

df3['N_Data'] = df3['D_Entrega'] + '!' + df3['Data']

df3 = df3.replace(to_replace='-', value='', regex=True)


df3['M_Data'] = df3['N_Data'].apply(func=lambda val: val[-11:] if len(val) > 11 else val)
df3 = df3.replace(to_replace='!', value='', regex=True)
print(df3)
