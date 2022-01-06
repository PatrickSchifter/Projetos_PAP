import requests
from requests.auth import HTTPBasicAuth
import pandas as pd

endpoint = 'http://homologacao.acomsistemas.com.br/api/adv/prenota/empresa/1?x-Pagina=1&data_inicial=2022%2F01%2F04&data_final=2022%2F01%2F05&x-Entidade=2020154'
entidade = '2020154'
data_i = '2022/01/03'
data_f = '2022/01/04'
user = 'PORTOAPORTO2022'
password = 'POR1355'
session = requests.Session()
session.auth = (user, password)
request = session.get(url=endpoint, headers={"Authorization": entidade})

for info in request.text.split(','):
    print(info)


