import requests
import pandas as pd
import os
import json

import simplejson

from Variaveis import meses_str, mes_atual, ano_atual

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
    if cont == 50:
        d = {
            "consulta": l_notas
        }
        h = {"Authorization": token, "Content-Type": "application/json"}
        payload = simplejson.dumps(d)

        request = requests.post(url=endpoint, data=payload, headers=h)
        rqts.append(request.text)
        cont = 0
        l_notas = []

d = {
    "consulta": l_notas
}
h = {"Authorization": token, "Content-Type": "application/json"}
payload = simplejson.dumps(d)

request = requests.post(url=endpoint, data=payload, headers=h)
rqts.append(request.text)
cont = 0
l_notas = []

for itens in rqts:
    for x in itens.split(','):
        x = x.replace('{', '')
        x = x.replace('}', '')
        x = x.replace('"', '')
        x = x.replace('df', '')
        x = x.replace('[', '')
        x = x.replace(']', '')
        print(x)
