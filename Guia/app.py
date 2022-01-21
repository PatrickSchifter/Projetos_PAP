from config import capacity_file, dados_estoque, guia_file
import pandas as pd
from pre_nota import df_item_2, df_item_99, df_item_7, df_item_16, df_item_20, df_item_100, df_item_115


df_cap = pd.read_excel(io=capacity_file, sheet_name='Capacidade')
df_cap.columns = ['Item', 'Descrição', 'Capacidade_Caixas', 'Capacidade_Unidade', 'Padrão', 'Excluir']
df_cap = df_cap[['Item', 'Capacidade_Caixas', 'Capacidade_Unidade', 'Padrão']]
df_cap = df_cap.apply(func=lambda val: pd.to_numeric(val, errors='coerce'))
df_cap['Item'] = df_cap['Item'].astype(dtype='str')
df_cap['Item'] = df_cap['Item'].fillna(0)

cabecalho = {"cd_empresa": '', "cd_deposito": '', "id_item": '', "id_embalagem": '', "cd_item": '', "ds_item": '',
             "cd_ncm": "", '': "UN", "qt_estoque": ''}
clean_carac = ['[', '{', '"', ']', '}']
list_dep = ['2', '7', '100', '115', '16', '20', '99']
list_info = []
pages_info = []
index = ['cd_item', 'ds_item', 'cd_deposito', 'qt_estoque']

for deposito in list_dep:
    for pagina in range(1, 8):  # range de 1 até 8
        pagina = str(pagina)
        pages_info.append(dados_estoque('1', deposito, pagina))

for infor in pages_info:
    for info in infor.split(','):
        for carac in clean_carac:
            info = info.replace(carac, '')
        if "cd_empresa" in info:
            y = info.split(':')
            list_info.append(cabecalho)
            cabecalho = {"cd_empresa": '', "cd_deposito": '', "id_item": '', "id_embalagem": '', "cd_item": '',
                         "ds_item": '', "cd_ncm": "", "sg_unidademedida": "UN", "qt_estoque": ''}
            cabecalho["cd_empresa"] = y[1]
        elif "cd_deposito" in info:
            y = info.split(':')
            cabecalho["cd_deposito"] = y[1]
        elif "id_item" in info:
            y = info.split(':')
            cabecalho["id_item"] = y[1]
        elif "id_embalagem" in info:
            y = info.split(':')
            cabecalho["id_embalagem"] = y[1]
        elif "cd_item" in info:
            y = info.split(':')
            try:
                y[1] = int(y[1])
                cabecalho["cd_item"] = y[1]
            except ValueError:
                cabecalho["cd_item"] = '-'
        elif "ds_item" in info:
            y = info.split(':')
            cabecalho["ds_item"] = y[1]
        elif "cd_ncm" in info:
            y = info.split(':')
            cabecalho["cd_ncm"] = y[1]
        elif "sg_unidademedida" in info:
            y = info.split(':')
            cabecalho["sg_unidademedida"] = y[1]
        elif "qt_estoque" in info:
            y = info.split(':')
            y[1] = y[1].split('.')
            cabecalho["qt_estoque"] = int(y[1][0])

df_estoque = pd.DataFrame(list_info)
df_estoque = df_estoque.query("qt_estoque != 0")
df_estoque = df_estoque[index]

df_2 = df_estoque.query("cd_deposito == '2'")
index2 = ['cd_item', 'qt_estoque']
df_2 = df_2[index2]
df_2.columns = ['Item', 'Show_Room_2']
df_2['Item'] = df_2['Item'].astype(dtype='str')
df_2 = pd.merge(left=df_2, right=df_item_2, on='Item', how='left')
df_2 = df_2.fillna(0)
df_2['Show_Room_2'] = df_2['Show_Room_2'].apply(func=lambda val: pd.to_numeric(val))


df_7 = df_estoque.query("cd_deposito == '7'")
index7 = ['cd_item', 'ds_item', 'qt_estoque']
df_7 = df_7[index7]
df_7.columns = ['Item', 'Descrição_Item', 'Veloz_7']
df_7['Item'] = df_7['Item'].astype(dtype='str')
df_7 = pd.merge(left=df_7, right=df_item_7, on='Item', how='left')
df_7 = df_7.fillna(0)
df_7['Veloz_7'] = df_7['Veloz_7'] - df_7['Quantidade']
df_7 = df_7[['Item', 'Descrição_Item', 'Veloz_7']]

df_100 = df_estoque.query("cd_deposito == '100'")
index100 = ['cd_item', 'qt_estoque']
df_100 = df_100[index100]
df_100.columns = ['Item', 'Caixas_Pardas_Veloz_100']
df_100['Item'] = df_100['Item'].astype(dtype='str')
df_100 = pd.merge(left=df_100, right=df_item_100, on='Item', how='left')
df_100 = df_100.fillna(0)
df_100['Caixas_Pardas_Veloz_100'] = df_100['Caixas_Pardas_Veloz_100'] - df_100['Quantidade']
df_100 = df_100[['Item', 'Caixas_Pardas_Veloz_100']]

df_115 = df_estoque.query("cd_deposito == '115'")
index115 = ['cd_item', 'qt_estoque']
df_115 = df_115[index115]
df_115.columns = ['Item', 'Bloqueios_115']
df_115['Item'] = df_115['Item'].astype(dtype='str')
df_115 = pd.merge(left=df_115, right=df_item_115, on='Item', how='left')
df_115 = df_115.fillna(0)
df_115['Bloqueios_115'] = df_115['Bloqueios_115'] - df_115['Quantidade']
df_115 = df_115[['Item', 'Bloqueios_115']]

df_16 = df_estoque.query("cd_deposito == '16'")
index16 = ['cd_item', 'qt_estoque']
df_16 = df_16[index16]
df_16.columns = ['Item', 'Reserva_Madero_16']
df_16['Item'] = df_16['Item'].astype(dtype='str')
df_16 = pd.merge(left=df_16, right=df_item_16, on='Item', how='left')
df_16 = df_16.fillna(0)
df_16['Reserva_Madero_16'] = df_16['Reserva_Madero_16'] - df_16['Quantidade']
df_16 = df_16[['Item', 'Reserva_Madero_16']]

df_20 = df_estoque.query("cd_deposito == '20'")
index20 = ['cd_item', 'qt_estoque']
df_20 = df_20[index20]
df_20.columns = ['Item', 'Avarias_20']
df_20['Item'] = df_20['Item'].astype(dtype='str')
df_20 = pd.merge(left=df_20, right=df_item_20, on='Item', how='left')
df_20 = df_20.fillna(0)
df_20['Reserva_Madero_16'] = df_20['Avarias_20'] - df_20['Quantidade']
df_20 = df_20[['Item', 'Avarias_20']]

df_99 = df_estoque.query("cd_deposito == '99'")
index99 = ['cd_item', 'qt_estoque']
df_99 = df_99[index99]
df_99.columns = ['Item', 'Bloqueios_Importação_99']
df_99['Item'] = df_99['Item'].astype(dtype='str')
df_99 = pd.merge(left=df_99, right=df_item_99, on='Item', how='left')
df_99 = df_99.fillna(0)
df_99['Bloqueios_Importação_99'] = df_99['Bloqueios_Importação_99'] - df_99['Quantidade']
df_99 = df_99[['Item', 'Bloqueios_Importação_99']]

df_guia = pd.merge(left=df_7, right=df_2, how='left', on='Item')
df_guia = pd.merge(left=df_guia, right=df_100, how='left', on='Item')
df_guia = pd.merge(left=df_guia, right=df_16, how='left', on='Item')
df_guia = pd.merge(left=df_guia, right=df_115, how='left', on='Item')
df_guia = pd.merge(left=df_guia, right=df_99, how='left', on='Item')
df_guia = pd.merge(left=df_guia, right=df_20, how='left', on='Item')
df_guia = pd.merge(left=df_guia, right=df_cap, how='left', on='Item')

df_guia['Sugestão_Unidades'] = df_guia['Capacidade_Unidade'] - df_guia['Show_Room_2']
df_guia['Sugestão_Caixas'] = df_guia['Sugestão_Unidades'] / df_guia['Padrão']
df_guia['Guia_Unidades'] = df_guia['Sugestão_Caixas'].apply(func=lambda val: '')
df_guia['Guia_Caixas'] = df_guia['Sugestão_Caixas'].apply(func=lambda val: '')
index_guia = ['Item', 'Descrição_Item', 'Veloz_7', 'Show_Room_2', 'Capacidade_Unidade', 'Sugestão_Unidades', 'Padrão',
              'Sugestão_Caixas', 'Guia_Unidades', 'Guia_Caixas', 'Caixas_Pardas_Veloz_100', 'Bloqueios_115',
              'Reserva_Madero_16', 'Avarias_20', 'Bloqueios_Importação_99']

df_guia = df_guia[index_guia]

writer = pd.ExcelWriter(path=guia_file)
df_guia.to_excel(excel_writer=writer, sheet_name='Sugestão_Guia', index=False)
writer.save()

print(df_estoque)
