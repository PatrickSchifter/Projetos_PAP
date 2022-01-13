
import pandas as pd
lista = []

item1 = {'nf': 'NaN', 'status': 'NaN', 'unidade': 'NaN', 'data': 'NaN'}
item1['nf'] = '6135'
print(item1)
lista.append(item1)
item1 = {'nf': '6281', 'status': 'ENTRADA', 'unidade': 'CO PINHAIS 02', 'data': '2022-01-04 20:49:01'}
lista.append(item1)
item1 = {'nf': '6327', 'status': 'CONTATE SEU FORNECEDOR(-104)', 'unidade': 'CO BELFORD ROXO 02', 'data': '2022-01-10 07:37:24'}
lista.append(item1)

df = pd.DataFrame(lista)

print(df)