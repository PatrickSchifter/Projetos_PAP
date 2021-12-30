import requests
import pandas as pd
import os
import simplejson as json

file_notas = 'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Aquivos Grande Adega/Notas/'

for itens in list(os.listdir(file_notas)):
    try:
        notas = pd.read_excel(file_notas + itens)
    except:
        break

cnpj = '26253785000106'
notas['Unnamed: 1'] = notas['Unnamed: 1'].astype(dtype='str')
token = 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJqdGkiOjEwMjgzNSwiZHQiOiIyMDIxMDMxNCJ9.M0ahvMZQ4-HOWDqFWe3Og05ZTeIhvQxkppcIWau1iKs '
endpoint = 'http://www.jadlog.com.br/embarcador/api/tracking/consultar'
nota_exemplo = "5996"

request = requests.post(url=endpoint, data={"consulta":
                                                [{"df":
                                                      {"nf": nota_exemplo, "serie": "001", "tpDocumento": 1}}]},
                                        headers={"Content-Type": "application/json", "Authorization": token})


print(request.text)
print(request.reason)
print(request.status_code)
