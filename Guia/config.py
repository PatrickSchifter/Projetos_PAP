import requests
import simplejson
from senhas import entidade, header
import time
from datetime import datetime, date

capacity_file = 'K:/CWB/Logistica/Rastreamento/Automacoes/Guia/CAPACIDADE DE ESTOQUE.xls'
guia_file = 'K:/CWB/Logistica/Rastreamento/Automacoes/Guia/Guia.xlsx'
header = simplejson.dumps(header)
header = simplejson.loads(header)

dias_meses = {
    1: 31,
    2: 28,
    3: 31,
    4: 30,
    5: 31,
    6: 30,
    7: 31,
    8: 31,
    9: 30,
    10: 31,
    11: 30,
    12: 31
}

def calc_data():
    dia_hoje = int(datetime.today().strftime('%d'))
    mes_hoje = int(datetime.today().strftime('%m'))
    ano_hoje = int(datetime.today().strftime('%Y'))

    data = datetime.weekday(datetime.today())

    if mes_hoje != 1:
        if data == 0:
            if dia_hoje == 1:
                dia = dias_meses[int(mes_hoje) - 1] - 3
                return date(ano_hoje, (mes_hoje - 1), dia).strftime("%d-%m-%Y")

            elif dia_hoje == 2:
                dia = dias_meses[int(mes_hoje) - 1] - 2
                return date(ano_hoje, int(mes_hoje) - 1, dia).strftime("%d-%m-%Y")

            elif dia_hoje == 3:
                dia = dias_meses[int(mes_hoje)] - 1
                return date(ano_hoje, int(mes_hoje) - 1, dia).strftime("%d-%m-%Y")

            else:
                dia = dia_hoje - 4
                return date(ano_hoje, int(mes_hoje), dia).strftime("%d-%m-%Y")
        else:
            if dia_hoje == 1:
                dia = dias_meses[int(mes_hoje) - 1] -1
                return date(ano_hoje, int(mes_hoje) - 1, dia).strftime("%d-%m-%Y")
            else:
                dia = dia_hoje - 2
                return date(ano_hoje, int(mes_hoje), dia).strftime("%d-%m-%Y")
    else:
        if data == 0:
            if dia_hoje == 1:
                dia = dias_meses[12] - 3
                return date(ano_hoje - 1, 12, dia).strftime("%d-%m-%Y")

            elif dia_hoje == 2:
                dia = dias_meses[12] - 3
                return date(ano_hoje - 1, 12, dia).strftime("%d-%m-%Y")

            elif dia_hoje == 3:
                dia = dias_meses[12] - 3
                return date(ano_hoje - 1, 12, dia).strftime("%d-%m-%Y")
            elif dia_hoje == 4:
                dia = dias_meses[12] - 3
                return date(ano_hoje - 1, 12, dia).strftime("%d-%m-%Y")

            else:
                dia = dia_hoje - 4
                return date(ano_hoje - 1, 12, dia).strftime("%d-%m-%Y")
        else:
            if dia_hoje == 1:
                dia = dias_meses[12] - 1
                return date(ano_hoje - 1, 12, dia).strftime("%d-%m-%Y")
            else:
                dia = dia_hoje - 2
                return date(ano_hoje, int(mes_hoje), dia).strftime("%d-%m-%Y")


def dados_estoque(codigo_empresa, deposito, pagina):
    endpoint = f'https://producao.acomsistemas.com.br/api/adm/estoque/{codigo_empresa}/deposito/{deposito}?x-Pagina={pagina}&x-Entidade={entidade}'
    request = requests.get(url=endpoint, headers=header, verify=False)
    time.sleep(1)
    if request.status_code == 200:
        return request.text
    else:
        time.sleep(1)
        return request.text

def dados_notas(codigo_empresa, pagina):
    data_inicial = calc_data()
    data_final = datetime.today()
    endpoint = f'https://producao.acomsistemas.com.br/api/adv/prenota/empresa/{codigo_empresa}?x-Pagina={pagina}&data_inicial={data_inicial[6:]}%2F{data_inicial[3:5]}%2F{data_inicial[:2]}&data_final={data_inicial[6:]}%2F{data_inicial[3:5]}%2F{data_inicial[6:]}&x-Entidade={entidade}'
    request = requests.get(url=endpoint, headers=header, verify=False)
    time.sleep(1)
    if request.status_code == 200:
        return request.text
    else:
        time.sleep(1)
        return request.text


# print(dados_estoque('1','2','1'))