import pandas as pd

df_resp = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Responsaveis.xlsx'

rep = 'VICTORIA REPRESENTACOES'

df_resp = pd.read_excel(io=df_resp, index_col=0)

print(df_resp.loc[rep]['email'])
