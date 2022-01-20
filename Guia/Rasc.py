import pandas as pd

dict = {'Coluna1': [25,23,34,45,34,65], 'Coluna2': [25,23,34,45,34,65]}



df = pd.DataFrame(data=dict)

df['Coluna3'] = df['Coluna1'] + df['Coluna2']
print(df)

