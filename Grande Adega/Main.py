import pandas as pd

final_file = 'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Aquivos Grande Adega/Prazos.txt'
df = pd.read_csv(final_file, ',')
print(df.columns)
# df = df.query("Lead_time_Jadlog != '-' & ")
print(df)
