from config import dados_notas
import pandas as pd

info_notas = []
clean_carac = ['[', '{', '"', ']', '}', 'qt_embalagem_faturada']
nota = {'cd_deposito': '', 'id_workflowestagio': '', 'itens': []}
data_notas = []


for page in range(1, 4):
    page = str(page)
    info_notas.append(dados_notas('1', page))


if len(info_notas) > 1:
    for infos in info_notas:
        #
        for info in infos.split(','):
            for clean in clean_carac:
                info = info.replace(clean, '')

            if 'cd_deposito' in info:
                data_notas.append(nota)
                info = info.split(':')
                nota = {'cd_deposito': '', 'id_workflowestagio': '', 'itens': []}
                nota['cd_deposito'] = info[1]

            elif 'id_workflowestagio' in info:
                info = info.split(':')
                nota['id_workflowestagio'] = info[1]

            elif 'id_workflowestagio' in info:
                info = info.split(':')
                nota['id_workflowestagio'] = info[1]

            elif 'nr_prenota' in info:
                print(info)

            elif 'qt_embalagem' in info:
                info = info.split(':')
                info[1] = info[1].split('.')
                info[1][0] = int(info[1][0])
                if info[1][0] > 0:
                    print(info[1][0])
                    nota['itens'].append(info[1])

            elif 'cd_item_nota' in info:
                info = info.split(':')
                print(info[1])
                nota['itens'].append(info[1])

df_itens = pd.DataFrame(data_notas)

print(df_itens)




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
