from datetime import datetime
from datetime import date

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

dia_atual = int(1)
mes_atual = int(1)
ano_atual = int(2022)


def dia_ant_ga():
    dia_semana = datetime.weekday()

    if dia_semana == 0:
        if mes_atual != 1:
            if dia_atual >= 4:
                if mes_atual < 10:
                    data_1 = '0' + str(dia_atual - 3) + '/0' + str(mes_atual) + '/' + str(ano_atual)
                    data_2 = '0' + str(dia_atual - 1) + '/0' + str(mes_atual) + '/' + str(ano_atual)
                else:
                    data_1 = '0' + str(dia_atual - 3) + '/' + str(mes_atual) + '/' + str(ano_atual)
                    data_2 = '0' + str(dia_atual - 1) + '/' + str(mes_atual) + '/' + str(ano_atual)
                ret = [data_1, data_2]
                return ret
            elif dia_atual == 3:
                if mes_atual < 10:
                    data_1 = str(dias_meses[mes_atual - 1]) + '/0' + str(mes_atual - 1) + '/' + str(ano_atual)
                    data_2 = '0' + str(dia_atual - 1) + '/0' + str(mes_atual) + '/' + str(ano_atual)
                elif mes_atual == 10:
                    data_1 = str(dias_meses[mes_atual - 1]) + '/0' + str(mes_atual - 1) + '/' + str(ano_atual)
                    data_2 = '0' + str(dia_atual - 1) + '/' + str(mes_atual) + '/' + str(ano_atual)
                else:
                    data_1 = str(dias_meses[mes_atual - 1]) + '/' + str(mes_atual - 1) + '/' + str(ano_atual)
                    data_2 = '0' + str(dia_atual - 1) + '/' + str(mes_atual) + '/' + str(ano_atual)
                ret = [data_1, data_2]
                return ret
            elif dia_atual == 2:
                if mes_atual < 10:
                    data_1 = str(dias_meses[mes_atual - 1] - 1) + '/0' + str(mes_atual - 1) + '/' + str(ano_atual)
                    data_2 = '0' + str(dia_atual - 1) + '/0' + str(mes_atual) + '/' + str(ano_atual)
                elif mes_atual == 10:
                    data_1 = str(dias_meses[mes_atual - 1] - 1) + '/0' + str(mes_atual - 1) + '/' + str(ano_atual)
                    data_2 = '0' + str(dia_atual - 1) + '/' + str(mes_atual) + '/' + str(ano_atual)
                else:
                    data_1 = str(dias_meses[mes_atual - 1] - 1) + '/' + str(mes_atual - 1) + '/' + str(ano_atual)
                    data_2 = '0' + str(dia_atual - 1) + '/' + str(mes_atual) + '/' + str(ano_atual)
                ret = [data_1, data_2]
                return ret
            elif dia_atual == 1:
                if mes_atual < 10:
                    data_1 = str(dias_meses[mes_atual - 1] - 2) + '/0' + str(mes_atual - 1) + '/' + str(ano_atual)
                    data_2 = str(dias_meses[mes_atual - 1]) + '/0' + str(mes_atual - 1) + '/' + str(ano_atual)
                elif mes_atual == 10:
                    data_1 = str(dias_meses[mes_atual - 1] - 2) + '/0' + str(mes_atual - 1) + '/' + str(ano_atual)
                    data_2 = str(dias_meses[mes_atual - 1]) + '/0' + str(mes_atual - 1) + '/' + str(ano_atual)
                else:
                    data_1 = str(dias_meses[mes_atual - 1] - 2) + '/' + str(mes_atual - 1) + '/' + str(ano_atual)
                    data_2 = str(dias_meses[mes_atual - 1]) + '/' + str(mes_atual - 1) + '/' + str(ano_atual)
                ret = [data_1, data_2]
                return ret
        else:  # Se o mês for Janeiro
            if dia_atual >= 4:
                data_1 = '0' + str(dia_atual - 3) + '/0' + str(mes_atual) + '/' + str(ano_atual)
                data_2 = '0' + str(dia_atual - 1) + '/0' + str(mes_atual) + '/' + str(ano_atual)
                ret = [data_1, data_2]
                return ret
            elif dia_atual == 3:
                data_1 = str(dias_meses[12]) + '/' + str(12) + '/' + str(ano_atual - 1)
                data_2 = '0' + str(dia_atual - 1) + '/0' + str(mes_atual) + '/' + str(ano_atual)
                ret = [data_1, data_2]
                return ret
            elif dia_atual == 2:
                data_1 = str(dias_meses[12] - 1) + '/' + str(12) + '/' + str(ano_atual - 1)
                data_2 = '0' + str(dia_atual - 1) + '/0' + str(mes_atual) + '/' + str(ano_atual)
                ret = [data_1, data_2]
                return ret
            elif dia_atual == 1:
                data_1 = str(dias_meses[12] - 2) + '/' + str(12) + '/' + str(ano_atual - 1)
                data_2 = str(dias_meses[12]) + '/' + str(12) + '/' + str(ano_atual - 1)
                ret = [data_1, data_2]
                return ret
    else:
        if mes_atual < 10:
            if dia_atual < 10:
                data_1 = '0' + str(dia_atual - 1) + '/0' + str(mes_atual) + '/' + str(ano_atual)
            else:
                data_1 = str(dia_atual - 1) + '/0' + str(mes_atual) + '/' + str(ano_atual)
            ret = [data_1, data_1]
            return ret
        else:
            if dia_atual < 10:
                data_1 = '0' + str(dia_atual - 1) + '/' + str(mes_atual) + '/' + str(ano_atual)
            else:
                data_1 = str(dia_atual - 1) + '/' + str(mes_atual) + '/' + str(ano_atual)
            ret = [data_1, data_1]
            return ret

print(dia_ant_ga())
