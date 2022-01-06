from datetime import datetime
import datetime
from datetime import date

import pandas as pd

dir_coletas = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Automação de Monitoramento/Coletados/'
dir_arq_coletas = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Automação de Monitoramento/Coletados/Arquivo/'
separador = "#" * 500
dia_hoje = int(datetime.date.today().strftime("%d"))
dia_semana = datetime.date.weekday(datetime.date.today())
ano_tual = int(datetime.date.today().strftime("%Y"))
dia_atual = date.today().strftime('%d')
ano_atual = date.today().strftime('%Y')
mes_atual = date.today().strftime('%m')
today = date.today().strftime("%d-%m-%Y")
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
meses = {"01": "janeiro",
         "02": "fevereiro",
         "03": "março",
         "04": "abril",
         "05": "maio",
         "06": "junho",
         "07": "julho",
         "08": "agosto",
         "09": "setembro",
         "10": "outubro",
         "11": "novembro",
         "12": "dezembro"}


def columns_to_datetime(val):
    val[11] = pd.to_datetime(arg=val[11], errors='coerce', format='%d-%m-%Y')
    print(val[11])
    val[12] = pd.to_datetime(arg=val[12], errors='coerce', format='%d-%m-%Y')
    print(val[12])
    val[13] = pd.to_datetime(arg=val[13], errors='coerce', format='%d-%m-%Y')
    val[14] = pd.to_datetime(arg=val[14], errors='coerce', format='%d-%m-%Y')
    return val[11], val[12], val[13], val[14]


def to_date_time(val):
    val = pd.to_datetime(arg=val, errors='coerce', format='%d-%m-%Y')
    valor = val
    return valor


def fatiamento(val):
    valor = '-'
    try:
        qtd = len(val)
        loc_sep = val.index('!')
        if qtd > 3:

            if loc_sep == 10:
                valor = val[:10]
                valor = valor.replace('/', '-')

            elif loc_sep == 1:
                valor = val[2:]
                valor = valor.replace('/', '-')
    except TypeError:
        print(f'Deu erro nesse valor {val} na hora do fatiamento verificar def fatiamento.')
        valor = '-'
    return valor


# Patrick, se isso der erro por conta de alguma incoerência nos dados, use split e depois compare os itens da lista.

def data():
    if mes_atual != '01':
        if dia_semana == 0:
            if dia_hoje == 1:
                dia = dias_meses[int(mes_atual) - 1] - 2
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d/%m/%Y")

            elif dia_hoje == 2:
                dia = dias_meses[int(mes_atual) - 1] - 1
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d/%m/%Y")

            elif dia_hoje == 3:
                dia = dias_meses[mes_atual]
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d/%m/%Y")

            else:
                dia = dia_hoje - 3
                return datetime.date(ano_tual, int(mes_atual), dia).strftime("%d/%m/%Y")

        else:
            if dia_hoje == 1:
                dia = dias_meses[int(mes_atual) - 1]
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d/%m/%Y")
            else:
                dia = dia_hoje - 1
                return datetime.date(ano_tual, int(mes_atual), dia).strftime("%d/%m/%Y")
    else:
        if dia_semana == 0:
            if dia_hoje == 1:
                dia = dias_meses[int(mes_atual) - 1] - 2
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d/%m/%Y")

            elif dia_hoje == 2:
                dia = dias_meses[int(mes_atual) - 1] - 1
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d/%m/%Y")

            elif dia_hoje == 3:
                dia = dias_meses[mes_atual]
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d/%m/%Y")

            else:
                dia = dia_hoje - 3
                return datetime.date(ano_tual, int(mes_atual), dia).strftime("%d/%m/%Y")

        else:
            if dia_hoje == 1:
                dia = dias_meses[int(mes_atual) - 1]
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d/%m/%Y")
            else:
                dia = dia_hoje - 1
                return datetime.date(ano_tual, int(mes_atual), dia).strftime("%d/%m/%Y")


def dataf():
    if mes_atual != '01':
        if dia_semana == 0:
            if dia_hoje == 1:
                dia = dias_meses[int(mes_atual) - 1] - 2
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d-%m-%Y")

            elif dia_hoje == 2:
                dia = dias_meses[int(mes_atual) - 1] - 1
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d-%m-%Y")

            elif dia_hoje == 3:
                dia = dias_meses[int(mes_atual)]
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d-%m-%Y")

            else:
                dia = dia_hoje - 3
                return datetime.date(ano_tual, int(mes_atual), dia).strftime("%d-%m-%Y")
        else:
            if dia_hoje == 1:
                dia = dias_meses[int(mes_atual) - 1]
                return datetime.date(ano_tual, int(mes_atual) - 1, dia).strftime("%d-%m-%Y")
            else:
                dia = dia_hoje - 1
                return datetime.date(ano_tual, int(mes_atual), dia).strftime("%d-%m-%Y")
    elif dia_semana == 0:
        if dia_hoje == 1:
            dia = dias_meses[12] - 2
            return datetime.date(ano_tual - 1, 12, dia).strftime("%d-%m-%Y")

        elif dia_hoje == 2:
            dia = dias_meses[12] - 2
            return datetime.date(ano_tual - 1, 12, dia).strftime("%d-%m-%Y")

        elif dia_hoje == 3:
            dia = dias_meses[12] - 2
            return datetime.date(ano_tual - 1, 12, dia).strftime("%d-%m-%Y")
        elif dia_hoje == 4:
            dia = dias_meses[12] - 2
            return datetime.date(ano_tual - 1, 12, dia).strftime("%d-%m-%Y")

        else:
            dia = dia_hoje - 3
            return datetime.date(ano_tual - 1, 12, dia).strftime("%d-%m-%Y")
    else:
        if dia_hoje == 1:
            dia = dias_meses[12]
            return datetime.date(ano_tual - 1, 12, dia).strftime("%d-%m-%Y")
        else:
            dia = dia_hoje - 1
            return datetime.date(ano_tual, int(mes_atual), dia).strftime("%d-%m-%Y")


dest_path = r'K:/CWB/Logistica/Rastreamento/Patrick/Storage/' + ano_atual + '/' + meses[mes_atual].title() + '/' + today
dia_ontem = int(date.today().strftime('%d')) - 1
ano_atual_i = int(date.today().strftime('%Y'))
yest = date(ano_atual_i, int(mes_atual), dia_ontem).strftime("%d-%m-%Y")

index_stat = ['Número', 'N.Pré-Nota', 'Emissão', 'Fantasia-Destinatário', 'Cidade-Destinatário', 'Uf',
              'Natureza-Fiscal', 'Situação-Fiscal', 'Descrição-Do-Depósito', 'Fantasia_Do_Transportador',
              'Fantasia-Comissionado', 'Data-De-Coleta', 'Previsão-Entrega', 'D_Entrega', 'Agendamento',
              'Lead-Time',
              'Dias-Para-Entrega', 'Resumo', 'Status']
path_r = r'K:/CWB/Logistica/Rastreamento/Patrick/Storage/'
partial_path = r'K:/CWB/Logistica/Rastreamento/Patrick/Storage/' + ano_atual + '/' + meses[mes_atual].title() + '/'
file = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/MONITORAMENTO ' + ano_atual + '.xlsx'
n_file = partial_path + today + '/Notas_emitidas ' + dataf() + '.xlsx'  # Notas do dia anterior
today = date.today().strftime("%d-%m-%Y")

path_dir_tod = partial_path + today

all_index = ['Número', 'N.Pré-Nota', 'Emissão', 'Fantasia-Destinatário', 'Cidade-Destinatário', 'Uf',
             'Natureza-Fiscal', 'Situação-Fiscal', 'Descrição-Do-Depósito', 'Fantasia_Do_Transportador',
             'Fantasia-Comissionado', 'Data-De-Coleta', 'Previsão-Entrega', 'D_Entrega', 'Agendamento',
             'Lead-Time',
             'Dias-Para-Entrega', 'Resumo', 'Observações']

indexes = ['Número', 'Data', 'Transportadora', 'Filial', 'Cidade/Estado', 'Ult_Status', 'Descrição']

aut_index = ['Número', 'N.Pré-Nota', 'Emissão', 'Fantasia-Destinatário', 'Cidade-Destinatário', 'Uf',
             'Natureza-Fiscal', 'Situação-Fiscal', 'Descrição-Do-Depósito', 'Fantasia_Do_Transportador',
             'Fantasia-Comissionado']

manual_index = ['Data-De-Coleta', 'Previsão-Entrega', 'D_Entrega', 'Agendamento', 'Lead-Time',
                'Dias-Para-Entrega', 'Resumo', 'Enviar-Email']

dest_file = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Automação de Monitoramento/MONITORAMENTO1 ' + ano_atual + '.xlsx'

meses_str = {"1": "JANEIRO",
             "2": "FEVEREIRO",
             "3": "MARÇO",
             "4": "ABRIL",
             "5": "MAIO",
             "6": "JUNHO",
             "7": "JULHO",
             "8": "AGOSTO",
             "9": "SETEMBRO",
             "10": "OUTUBRO",
             "11": "NOVEMBRO",
             "12": "DEZEMBRO"}

from datetime import date

itens_ign = ['tracking; ', 'success; ', 'message; ', 'header; ', 'remetente; ', 'destinatario; ', 'items; ',
             'item; ',
             'data_hora; ', 'dominio; ', 'filial; ', 'cidade; ', 'ocorrencia; ', 'descricao; ', 'tipo; ',
             'data_hora_efetiva; ', 'nome_recebedor; ', 'nro_doc_recebedor; ',
             '?xml version="1.0" encoding="UTF-8" ?; ', '']


def calc_data(dia, mes, ano, valor_a_somar):
    uteis = [0, 1, 2, 3, 4]
    if date.weekday(datetime.date(ano, mes, dia)) == 4 and valor_a_somar == 1:
        dia += 3
        valor_a_somar -= 1

    while valor_a_somar > 0:

        data = date.weekday(datetime.date(ano, mes, dia))

        if data in uteis:
            if mes == 12 and dias_meses[mes] == dia:
                dia = 1
                mes = 1
                ano += 1
                valor_a_somar -= 1
            elif dias_meses[mes] == dia:
                dia = 1
                mes += 1
                valor_a_somar -= 1
            elif data == 4:
                dia += 3
                valor_a_somar -= 1
            else:
                dia += 1
                valor_a_somar -= 1
        elif data == 5 and valor_a_somar == 1:
            dia += 2
            valor_a_somar -= 1

        else:
            dia += 1
    if dia < 10:
        dia = '0' + str(dia)
    else:
        dia = str(dia)
    if mes < 10:
        mes = '0' + str(mes)
    else:
        mes = str(mes)

    valor = dia + '-' + mes + '-' + str(ano)

    return valor


def func(val):
    val_ignore = ['', '-']
    val = val.split('!')
    try:
        if val[0] not in val_ignore and val[1] not in val_ignore:
            val[1] = int(val[1][:-2])
            dia = int(val[0][8:10])
            mes = int(val[0][5:7])
            ano = int(val[0][:4])
            val = calc_data(dia, mes, ano, val[1])
        else:
            val = '-'
    except ValueError:
        if val[0] not in val_ignore and val[1] not in val_ignore:
            try:
                val[1] = int(val[1][0])
            except TypeError:
                pass
            dia = int(val[0][:2])
            mes = int(val[0][3:5])
            ano = int(val[0][6:])
            val = calc_data(dia, mes, ano, val[1])
        else:
            val = '-'

    return val


full_name_urano = "Urano - " + str(date.today().strftime('%d-%m-%Y')) + ".CSV"

p_source_path = "C:/Users/patrick.paula/Downloads/"

source_path_urano = p_source_path + full_name_urano
dest_path_urano = dest_path + "/" + full_name_urano

column_veloz = ['Faturamento', 'Data', 'Expedição', 'Número', 'Pre_nota', 'Volumes', 'Transportadora', 'NF_Veloz',
                'Status', 'Data_De_Coleta', 'Observação', 'Atual', 'Pendencia', 'Aguardando_Separação', 'Nao_coletados'
                ]


def conversor_dt(val):
    valor = '-'
    if len(val) > 1:
        valor = val[8:10] + '/' + val[5:7] + '/' + val[:4]
    return valor


def conversor_ldt(val):
    valor = '-'
    if val == '-':
        valor = '-'
    else:
        try:
            valor = val[:-2]
        except ValueError:
            if val.isnumeric():
                valor = str(val)
            else:
                valor = '-'

    return valor


def calc_dias_p_entrega(dia_prev, mes_prev, ano_prev):
    if dia_prev == int(dia_atual):
        cont = 0
        return cont if cont >= 0 else f'Em atraso ({cont})'
    elif dia_prev < int(dia_atual) and mes_prev <= int(mes_atual) and ano_prev <= int(ano_atual):
        uteis = [0, 1, 2, 3, 4]
        cont = 0
        t_loop = True
        while t_loop:
            try:
                data = date.weekday(datetime.date(ano_prev, mes_prev, dia_prev))
            except:
                pass
            if dia_prev > 0:
                if data in uteis:
                    dia_prev += 1
                    cont -= 1
                else:
                    dia_prev += 1
            elif dia_prev == 0 and mes_prev != 12:
                mes_prev = mes_prev + 1
                dia_prev = dias_meses[mes_prev + 1]
                cont -= 1
            else:
                mes_prev = 1
                dia_prev = dias_meses[mes_prev]
                ano_prev = ano_prev + 1
                cont -= 1
            if dia_prev == int(dia_atual):
                if mes_prev == int(mes_atual):
                    if ano_prev == int(ano_atual):
                        t_loop = False
                        return cont if cont >= 0 else f'Em atraso ({cont})'
    else:
        uteis = [0, 1, 2, 3, 4]
        cont = 0
        t_loop = True
        while t_loop:
            try:
                data = date.weekday(datetime.date(ano_prev, mes_prev, dia_prev))
            except:
                pass
            if dia_prev > 0:
                if data in uteis:
                    dia_prev -= 1
                    cont += 1
                else:
                    dia_prev -= 1
            elif dia_prev == 0 and mes_prev != 1:
                mes_prev = mes_prev - 1
                dia_prev = dias_meses[mes_prev - 1]
                cont += 1
            else:
                mes_prev = 12
                dia_prev = dias_meses[mes_prev]
                ano_prev = ano_prev - 1
                cont += 1
            if dia_prev == int(dia_atual):
                if mes_prev == int(mes_atual):
                    if ano_prev == int(ano_atual):
                        t_loop = False
        return cont if cont >= 0 else f'Em atraso ({cont})'


def dias_p_entrega(val):
    valor = '-'
    l_conc = val.split('!')
    l_prev = l_conc[0]
    l_prev = l_prev.split('-')
    if l_conc[1] != '-':
        valor = "Entregue"
    elif l_conc[0] == '-' or l_conc[2] == '-':
        valor = '-'
    else:
        dia_prev = int(l_prev[0])
        mes_prev = int(l_prev[1])
        ano_prev = int(l_prev[2])
        valor = calc_dias_p_entrega(dia_prev, mes_prev, ano_prev)

    return valor
