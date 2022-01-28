from datetime import datetime, date

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

dia_hoje = 17
dia_semana = 0
ano_tual = 2022
dia_atual = '17'
ano_atual = '2022'
mes_atual = '01'

def dataf():
    if mes_atual != '01':
        if dia_semana == 0:
            if dia_hoje == 1:
                dia = dias_meses[int(mes_atual) - 1] - 2
                return date(ano_tual, int(mes_atual) - 1, dia).strftime("%d-%m-%Y")

            elif dia_hoje == 2:
                dia = dias_meses[int(mes_atual) - 1] - 1
                return date(ano_tual, int(mes_atual) - 1, dia).strftime("%d-%m-%Y")

            elif dia_hoje == 3:
                dia = dias_meses[int(mes_atual)]
                return date(ano_tual, int(mes_atual) - 1, dia).strftime("%d-%m-%Y")

            else:
                dia = dia_hoje - 3
                return date(ano_tual, int(mes_atual), dia).strftime("%d-%m-%Y")
        else:
            if dia_hoje == 1:
                dia = dias_meses[int(mes_atual) - 1]
                return date(ano_tual, int(mes_atual) - 1, dia).strftime("%d-%m-%Y")
            else:
                dia = dia_hoje - 1
                return date(ano_tual, int(mes_atual), dia).strftime("%d-%m-%Y")
    elif dia_semana == 0:
        if dia_hoje == 1:
            dia = dias_meses[12] - 2
            return date(ano_tual - 1, 12, dia).strftime("%d-%m-%Y")

        elif dia_hoje == 2:
            dia = dias_meses[12] - 2
            return date(ano_tual - 1, 12, dia).strftime("%d-%m-%Y")

        elif dia_hoje == 3:
            dia = dias_meses[12] - 2
            return date(ano_tual - 1, 12, dia).strftime("%d-%m-%Y")
        elif dia_hoje == 4:
            dia = dias_meses[12] - 2
            return date(ano_tual - 1, 12, dia).strftime("%d-%m-%Y")

        else:
            dia = dia_hoje - 3
            return date(ano_tual, int(mes_atual), dia).strftime("%d-%m-%Y")
    else:
        if dia_hoje == 1:
            dia = dias_meses[12]
            return date(ano_tual - 1, 12, dia).strftime("%d-%m-%Y")
        else:
            dia = dia_hoje - 1
            return date(ano_tual, int(mes_atual), dia).strftime("%d-%m-%Y")

print(dataf())