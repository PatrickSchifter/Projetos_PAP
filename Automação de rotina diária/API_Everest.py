import requests
import simplejson
import certifi
import pandas as pd

endpoint_pr = 'https://producao.acomsistemas.com.br/api/dfi/notaemitida?emissao_inicial=2022%2F01%2F17&emissao_final=2022%2F01%2F17&x-Entidade=2020005&x-Pagina=1'
endpoint = 'http://homologacao.acomsistemas.com.br:80/api/adv/prenota/empresa/1?x-Pagina=1&data_inicial=2022%2F01%2F10&data_final=2022%2F01%2F10&x-Entidade=2020154'

entidade_pr = '2020005'
entidade = '2020154'
user = 'PORTOAPORTO2022'
password = 'POR1355'

header = {'Authorization': 'Basic UE9SVE9BUE9SVE8yMDIyOlBPUjEzNTU='}
header = simplejson.dumps(header)
header = simplejson.loads(header)

request = requests.get(url=endpoint_pr, headers=header, verify=False)
list_to_consider = ["nr_prenota", "dh_emissao", "cd_deposito", "cd_transportador", "Código do destinatário"]

print(request.content)

# df = pd.DataFrame(request.text.split('['))
# print(df.columns)

