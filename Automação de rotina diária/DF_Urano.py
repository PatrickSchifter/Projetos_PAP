import pandas as pd
from Variaveis import dest_path_urano

df_urano = pd.read_csv(filepath_or_buffer=dest_path_urano, sep=';')
df_urano.fillna('-')
df_urano = df_urano.rename(columns={'Numero NF-e': 'Número', 'Data de emissao CT-e': "Emissão_cte",
                                    'Descricao Ocorrencia CT-e': 'Ocorrência',
                                    'Data de movimento Ocorrencia CT-e': 'Data_Ocorrência'})
df_urano = df_urano[['Número', 'Emissão_cte', 'Ocorrência', 'Data_Ocorrência']]

df_urano['Emissão_cte'] = df_urano['Emissão_cte'].astype(dtype='str')  # Conversão para fatiamento
df_urano['Data_Ocorrência'] = df_urano['Data_Ocorrência'].astype(dtype='str')  # Conversão para fatiamento
df_urano['Emissão_cte'] = df_urano['Emissão_cte'].apply(
    func=lambda val: val[:10])  # Fatiamento para retirar horas do datetime
df_urano['Data_Ocorrência'] = df_urano['Data_Ocorrência'].apply(
    func=lambda val: val[:10])  # Fatiamento para retirar horas do datetime
qy_ent = df_urano.query("Ocorrência == 'Entrega normal.'")
qy_ent = qy_ent.rename(columns={"Emissão_cte": "Emissão1",
                                'Ocorrência': 'Ocorrência1', 'Data_Ocorrência': 'Data_Ocorrência1'})
