import requests

url = 'https://ssw.inf.br/api/tracking'
cnpj = '00069957000194'

request = requests.post(url=url, data={'cnpj': cnpj, 'nro_nf': 417461})

print(request.text)
