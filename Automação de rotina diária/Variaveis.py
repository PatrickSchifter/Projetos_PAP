from datetime import datetime
import datetime
from datetime import date

dia_atual = date.today().strftime('%d')
ano_atual = date.today().strftime('%Y')
mes_atual = date.today().strftime('%m')

today = date.today().strftime("%d-%m-%Y")
partial_path = "K:/CWB/Logistica/Rastreamento/Patrick/Storage/"
path_dir_tod = partial_path + today

all_index = ['Número', 'N.Pré-Nota', 'Emissão', 'Fantasia-Destinatário', 'Cidade-Destinatário', 'Uf',
             'Natureza-Fiscal', 'Situação-Fiscal', 'Descrição-Do-Depósito', 'Fantasia_Do_Transportador',
             'Fantasia-Comissionado', 'Data-De-Coleta', 'Previsão-Entrega', 'D_Entrega', 'Agendamento',
             'Lead-Time',
             'Dias-Para-Entrega', 'Resumo', 'Observações']

indexes = ['Número', 'Data', 'Transportadora', 'Filial', 'Cidade/Estado', 'Ult_Status']

dest_file = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Automação de Monitoramento/MONITORAMENTO ' + ano_atual + '.xlsx'

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

dia_hoje = int(datetime.date.today().strftime("%d"))
dia_semana = datetime.date.weekday(datetime.date.today())
ano_tual = int(datetime.date.today().strftime("%Y"))


def data():
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

from datetime import date

today = date.today().strftime("%d-%m-%Y")
dest_path = "K:/CWB/Logistica/Rastreamento/Patrick/Storage/" + today
dia_ontem = int(date.today().strftime('%d')) - 1
ano_atual_i = int(date.today().strftime('%Y'))
yest = date(ano_atual_i, int(mes_atual), dia_ontem).strftime("%d-%m-%Y")

itens_ign = ['tracking; ', 'success; ', 'message; ', 'header; ', 'remetente; ', 'destinatario; ', 'items; ',
             'item; ',
             'data_hora; ', 'dominio; ', 'filial; ', 'cidade; ', 'ocorrencia; ', 'descricao; ', 'tipo; ',
             'data_hora_efetiva; ', 'nome_recebedor; ', 'nro_doc_recebedor; ',
             '?xml version="1.0" encoding="UTF-8" ?; ', '']
