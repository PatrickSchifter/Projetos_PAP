import pandas as pd
from Tratamento_dados import final_file

df = pd.read_csv(final_file, ',')
print(df.columns)
df = df.query("Lead_time_Jadlog != '-' & ")
print(df)
