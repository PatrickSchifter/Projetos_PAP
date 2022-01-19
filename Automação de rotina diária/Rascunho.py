import pandas as pd

str = '03-01-2022'

str = pd.to_datetime(arg=str, format='%d-%m-%Y')

print(str)