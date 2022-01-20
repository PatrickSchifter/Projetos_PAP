from config import dados_notas
import pandas as pd

info_notas = []
faturados = ['3', '6']
clean_carac = ['[', '{', '"', ']', '}', 'qt_embalagem_faturada', 'qt_embalagem_cortada']
nota = {'cd_deposito': '', 'id_workflowestagio': '', 'itens': []}
data_notas = []

for page in range(1, 8):
    page = str(page)
    info_notas.append(dados_notas('1', page))

if len(info_notas) > 1:
    for infos in info_notas:
        #
        for info in infos.split(','):
            for clean in clean_carac:
                info = info.replace(clean, '')

            if 'nr_prenota' in info:
                print(info)

            if 'cd_deposito' in info:
                data_notas.append(nota)
                info = info.split(':')
                nota = {'cd_deposito': '', 'id_workflowestagio': '', 'itens': []}
                nota['cd_deposito'] = info[1]
                print(info)

            elif 'id_workflowestagio' in info:
                info = info.split(':')
                nota['id_workflowestagio'] = info[1]
                print(info)

            elif 'id_workflowestagio' in info:
                info = info.split(':')
                nota['id_workflowestagio'] = info[1]
                print(info)

            elif 'qt_embalagem' in info:
                info = info.split(':')
                info[1] = info[1].split('.')
                info[1][0] = int(info[1][0])
                if info[1][0] > 0:
                    nota['itens'].append(info[1][0])
                print(info)

            elif 'cd_item_nota' in info:
                info = info.split(':')
                nota['itens'].append(info[1])
                print(info)

df_itens = pd.DataFrame(data_notas)


df_itens = df_itens.query("id_workflowestagio != '3' & id_workflowestagio != '6'")  # Filtro de não faturados

"""
Faturamento - 4
Bloqueado - 1
Liberado - 2
Separação - 5
Faturado - 6
Faturado - 3
Faturamento Veloz - 7
Cliente Retira - 8
"""


df_itens2 = df_itens = df_itens.query("cd_deposito == '2'")
df_itens7 = df_itens = df_itens.query("cd_deposito == '7'")
df_itens100 = df_itens = df_itens.query("cd_deposito == '100'")
df_itens115 = df_itens = df_itens.query("cd_deposito == '115'")
df_itens16 = df_itens = df_itens.query("cd_deposito == '16'")
df_itens20 = df_itens = df_itens.query("cd_deposito == '20'")
df_itens99 = df_itens = df_itens.query("cd_deposito == '99'")

dict_2_df = {'itens': [], 'qtds': []}
dict_7_df = {'itens': [], 'qtds': []}
dict_100_df = {'itens': [], 'qtds': []}
dict_115_df = {'itens': [], 'qtds': []}
dict_16_df = {'itens': [], 'qtds': []}
dict_20_df = {'itens': [], 'qtds': []}
dict_99_df = {'itens': [], 'qtds': []}

dict_2 = {}
dict_7 = {}
dict_100 = {}
dict_115 = {}
dict_16 = {}
dict_20 = {}
dict_99 = {}

def itens(list, dep):
    qtd_atual = 0
    nt_atual = ''
    for item in list:
        if type(item) is str:
            nt_atual = item
        elif type(item) is int:
            qtd_atual = item
        if qtd_atual != 0 and nt_atual != '':
            if dep == '2':
                if dict_2.get(nt_atual):
                    dict_2[nt_atual] += qtd_atual
                else:
                    dict_2[nt_atual] = qtd_atual
                qtd_atual = 0
                nt_atual = ''
            elif dep == '7':
                if dict_7.get(nt_atual):
                    dict_7[nt_atual] += qtd_atual
                else:
                    dict_7[nt_atual] = qtd_atual
                qtd_atual = 0
                nt_atual = ''
            elif dep == '100':
                if dict_100.get(nt_atual):
                    dict_100[nt_atual] += qtd_atual
                else:
                    dict_100[nt_atual] = qtd_atual
                qtd_atual = 0
                nt_atual = ''
            elif dep == '115':
                if dict_115.get(nt_atual):
                    dict_115[nt_atual] += qtd_atual
                else:
                    dict_115[nt_atual] = qtd_atual
                qtd_atual = 0
                nt_atual = ''
            elif dep == '16':
                if dict_16.get(nt_atual):
                    dict_16[nt_atual] += qtd_atual
                else:
                    dict_16[nt_atual] = qtd_atual
                qtd_atual = 0
                nt_atual = ''
            elif dep == '20':
                if dict_20.get(nt_atual):
                    dict_20[nt_atual] += qtd_atual
                else:
                    dict_20[nt_atual] = qtd_atual
                qtd_atual = 0
                nt_atual = ''
            elif dep == '99':
                if dict_99.get(nt_atual):
                    dict_99[nt_atual] += qtd_atual
                else:
                    dict_99[nt_atual] = qtd_atual
                qtd_atual = 0
                nt_atual = ''

df_itens2['itens'] = df_itens2['itens'].apply(func=lambda val: itens(val, '2'))
df_itens7['itens'] = df_itens7['itens'].apply(func=lambda val: itens(val, '7'))
df_itens100['itens'] = df_itens100['itens'].apply(func=lambda val: itens(val, '100'))
df_itens115['itens'] = df_itens115['itens'].apply(func=lambda val: itens(val, '115'))
df_itens16['itens'] = df_itens16['itens'].apply(func=lambda val: itens(val, '16'))
df_itens20['itens'] = df_itens20['itens'].apply(func=lambda val: itens(val, '20'))
df_itens99['itens'] = df_itens99['itens'].apply(func=lambda val: itens(val, '99'))


for item in dict_2:
    dict_2_df['itens'].append(item)
    dict_2_df['qtds'].append(dict_2[item])

for item in dict_7:
    dict_7_df['itens'].append(item)
    dict_7_df['qtds'].append(dict_7[item])

for item in dict_100:
    dict_100_df['itens'].append(item)
    dict_100_df['qtds'].append(dict_100[item])

for item in dict_115:
    dict_115_df['itens'].append(item)
    dict_115_df['qtds'].append(dict_115[item])

for item in dict_16:
    dict_16_df['itens'].append(item)
    dict_16_df['qtds'].append(dict_16[item])

for item in dict_20:
    dict_20_df['itens'].append(item)
    dict_20_df['qtds'].append(dict_20[item])

for item in dict_99:
    dict_99_df['itens'].append(item)
    dict_99_df['qtds'].append(dict_99[item])

df_item_2 = pd.DataFrame(dict_2_df)

writer = pd.ExcelWriter(path='C:/Users/patrick.paula/Desktop/Testes/Teste_guia2.xlsx')
df_item_2.to_excel(excel_writer=writer, sheet_name='Sugestão_Guia', index=False)
writer.save()


df_item_2.columns = ['Item', 'Quantidade']
df_item_7 = pd.DataFrame(dict_7_df)
df_item_7.columns = ['Item', 'Quantidade']
df_item_100 = pd.DataFrame(dict_100_df)
df_item_100.columns = ['Item', 'Quantidade']
df_item_115 = pd.DataFrame(dict_115_df)
df_item_115.columns = ['Item', 'Quantidade']
df_item_16 = pd.DataFrame(dict_16_df)
df_item_16.columns = ['Item', 'Quantidade']
df_item_20 = pd.DataFrame(dict_20_df)
df_item_20.columns = ['Item', 'Quantidade']
df_item_99 = pd.DataFrame(dict_99_df)
df_item_99.columns = ['Item', 'Quantidade']

