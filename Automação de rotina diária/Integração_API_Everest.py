import requests
from requests.auth import HTTPBasicAuth
import pandas as pd

endpoint = 'http://homologacao.acomsistemas.com.br/'
cnpj = '00069957000194'
entidade = '2020154'
data_i = '2022/01/03'
data_f = '2022/01/04'
usuario = 'PORTOAPORTO2022'
senha = 'POR1355'
usuario_senha = (usuario, senha)
request = requests.get(url=endpoint,
                       headers={"Authorization": entidade, 'x-Pagina': '1', 'data_inicial': data_i, 'data_final': data_f, 'x-Entidade': entidade,
                             'cd_empresa': '1'})
print(request.text)
print(request.reason)
print(request.status_code)
print(request.headers)
