import requests
import simplejson
import certifi
import pandas as pd
from Projetos.Rotina.config import endpoint_everest as endpoint_pr
from Projetos.Rotina.config import header_everest as header

header = simplejson.dumps(header)
header = simplejson.loads(header)

request = requests.get(url=endpoint_pr, headers=header, verify=False)
list_to_consider = ["nr_prenota", "dh_emissao", "cd_deposito", "cd_transportador", "Código do destinatário"]

print(request.content)

# df = pd.DataFrame(request.text.split('['))
# print(df.columns)
