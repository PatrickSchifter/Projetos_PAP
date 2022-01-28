import requests
import pandas as pd
from config import conversor_dt, to_date_time
import time
import simplejson

from config import meses_str, mes_atual, ano_atual, path_dir_tod

mes_atual = str(int(mes_atual))

sheet_m = meses_str[(mes_atual)] + '-' + ano_atual
try:
    sheet_m_a = meses_str[str(int(mes_atual) - 1)]
except KeyError:
    pass

file_ga = 'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/MONITORAMENTO GA 2022.xlsx'
df_ga = pd.read_excel(io=file_ga, sheet_name=sheet_m)
df_ga = df_ga.fillna('-')
df_ga = df_ga.query("D_Entrega == '-'")

token = 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJqdGkiOjEwMjgzNSwiZHQiOiIyMDIxMDMxNCJ9.M0ahvMZQ4-HOWDqFWe3Og05ZTeIhvQxkppcIWau1iKs'
endpoint = 'http://www.jadlog.com.br/embarcador/api/tracking/consultar'

cnpj = "26253785000106"

l_notas = []
rqts = []


def notas_f(nota):
    nota = str(nota)
    nf_search = {"df":
                     {"nf": nota,
                      "cnpjRemetente": "26253785000106",
                      "tpDocumento": 1}}
    return nf_search


cont = 0
for nota in df_ga['Número']:
    l_notas.append(notas_f(nota))
    cont += 1
    if cont == 40:
        d = {
            "consulta": l_notas
        }
        h = {"Authorization": token, "Content-Type": "application/json"}
        payload = simplejson.dumps(d)

        request = requests.post(url=endpoint, data=payload, headers=h)
        rqts.append(request.text)
        cont = 0
        l_notas = []
        time.sleep(1)

d = {
    "consulta": l_notas
}
h = {"Authorization": token, "Content-Type": "application/json"}
payload = simplejson.dumps(d)

request = requests.post(url=endpoint, data=payload, headers=h)
rqts.append(request.text)
cont = 0
l_notas = []

list_to_ig = ['recebedor', 'previsaoEntrega', 'eventos', 'shipmentId', 'dacte', 'tracking', 'valor', 'peso', 'dtEmissao', 'altura', 'largura', 'comprimento']
list_info = []
info = {'nf': 'NaN', 'status': 'NaN', 'unidade': 'NaN', 'data': 'NaN'}

for itens in rqts:
    for x in itens.split(','):
        x = x.replace('{', '')
        x = x.replace('}', '')
        x = x.replace('"', '')
        x = x.replace('df', '')
        x = x.replace('[', '')
        x = x.replace(']', '')
        x = x.replace(']', '')
        x = x.replace('consulta', '')
        x = x.replace('tpDocumento:1', '')
        x = x.replace('cnpjRemetente:26253785000106', '')
        x = x.replace('error:id:-1', '')
        x = x.replace('descricao:Nao localizado.', '')
        try:
            if x[0] == ':':
                x = x[-7:]
        except IndexError:
            pass
        for itens_to_ignore in list_to_ig:
            if itens_to_ignore in x:
                x = x.replace(x, '')

        if x != '':
            if 'nf' in x:
                y = x.split(':')
                list_info.append(info)
                info = {'nf': '-', 'status': '-', 'unidade': '-', 'data': '-'}
                info['nf'] = y[1]
            elif 'status' in x:
                y = x.split(':')
                info['status'] = y[1]
            elif 'unidade' in x:
                y = x.split(':')
                info['unidade'] = y[1]
            elif 'data' in x:
                y = x.split(':')
                info['data'] = y[1]
            else:
                continue

df_jad = pd.DataFrame(list_info)
df_jad['data'] = df_jad['data'].apply(func=lambda val: conversor_dt(val))
df_jad['data'] = df_jad['data'].apply(func=lambda val: val if len(val) > 5 else '-')
df_jad['data'] = df_jad['data'].apply(func=lambda val: to_date_time(val))
df_jad['nf'] = df_jad['nf'].apply(func=lambda val: pd.to_numeric(arg=val, errors='coerce'))
print(df_jad)
writer = pd.ExcelWriter(path_dir_tod + '/Relatório Jadlog.xlsx')
df_jad.to_excel(excel_writer=writer, index=False)
writer.save()