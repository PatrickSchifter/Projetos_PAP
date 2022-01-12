from datetime import date
from datetime import datetime

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

dia_hoje = int(datetime.today().strftime('%d'))
mes_hoje = int(datetime.today().strftime('%m'))
ano_hoje = int(datetime.today().strftime('%Y'))

def calc_dat_15():
    if dia_hoje > 15:
        dia = dia_hoje - 15
        mes = mes_hoje
        ano = ano_hoje
    else:
        try:
            dia = (dia_hoje - 15) + dias_meses[mes_hoje-1]
        except KeyError:
            pass
        if mes_hoje - 1 == 0:
            try:
                dia = (dia_hoje - 15) + dias_meses[12]
            except KeyError:
                pass
            mes = 12
            ano = ano_hoje -1
        else:
            try:
                mes = mes_hoje - 1
            except KeyError:
                pass
            ano = ano_hoje
    return str(dia) + '/' + str(mes) + '/' + str(ano)

print(type(dia_hoje))