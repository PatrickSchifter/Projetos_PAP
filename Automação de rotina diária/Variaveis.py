from datetime import datetime
import datetime
from datetime import date

separador = "#" * 500
dia_hoje = int(datetime.date.today().strftime("%d"))
dia_semana = datetime.date.weekday(datetime.date.today())
ano_tual = int(datetime.date.today().strftime("%Y"))


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


dia_atual = date.today().strftime('%d')
ano_atual = date.today().strftime('%Y')
mes_atual = date.today().strftime('%m')
today = date.today().strftime("%d-%m-%Y")
dest_path = "K:/CWB/Logistica/Rastreamento/Patrick/Storage/" + today
dia_ontem = int(date.today().strftime('%d')) - 1
ano_atual_i = int(date.today().strftime('%Y'))
yest = date(ano_atual_i, int(mes_atual), dia_ontem).strftime("%d-%m-%Y")

index_stat = ['Número', 'N.Pré-Nota', 'Emissão', 'Fantasia-Destinatário', 'Cidade-Destinatário', 'Uf',
              'Natureza-Fiscal', 'Situação-Fiscal', 'Descrição-Do-Depósito', 'Fantasia_Do_Transportador',
              'Fantasia-Comissionado', 'Data-De-Coleta', 'Previsão-Entrega', 'D_Entrega', 'Agendamento',
              'Lead-Time',
              'Dias-Para-Entrega', 'Resumo', 'Status']

partial_path = "K:/CWB/Logistica/Rastreamento/Patrick/Storage/"
file = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/MONITORAMENTO 2021 v.1.2.xlsx'
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


full_name_urano = "Urano " + str(date.today().strftime('%d-%m-%Y')) + ".CSV"

p_source_path = "C:/Users/patrick.paula/Downloads/"
dest_path = "K:/CWB/Logistica/Rastreamento/Patrick/Storage/" + today

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
