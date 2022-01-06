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

notas['Unnamed: 1'] = notas['Unnamed: 1'].astype(dtype='str')
token = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJqdGkiOjEwMjgzNSwiZHQiOiIyMDIxMDMxNCJ9.M0ahvMZQ4-HOWDqFWe3Og05ZTeIhvQxkppcIWau1iKs'
endpoint = 'http://www.jadlog.com.br/embarcador/api/tracking/consultar'
nota_exemplo = "5996"

user = '26253785000106'
password = '4DzYlSX'
codigo_cliente = '102835'
cd_franqui = '1766'

session = requests.Session()
session.auth = (user, password)
request = session.post(url=endpoint, data={"consulta":
                                                [{"df":
                                                      {"nf": nota_exemplo, "tpDocumento": 1}}]},
                                        headers={"Authorization": token, "Content-Type": "application/json"})



print(request.text)
print(request.reason)
print(request.status_code)
