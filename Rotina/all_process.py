import pandas as pd
from datetime import datetime
from datetime import date


class Process:
    def __init__(self):
        dir_coletas = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Automação de Monitoramento/Coletados/'
        dir_arq_coletas = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Automação de Monitoramento/Coletados/Arquivo/'
        separador = "#" * 500
        dia_hoje = int(date.today().strftime("%d"))
        dia_semana = date.weekday(date.today())
        ano_tual = int(date.today().strftime("%Y"))
        dia_atual = date.today().strftime('%d')
        ano_atual = date.today().strftime('%Y')
        mes_atual = date.today().strftime('%m')
        today = date.today().strftime("%d-%m-%Y")
        self.dias_meses = {
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

        dest_path = r'K:/CWB/Logistica/Rastreamento/Patrick/Storage/' + ano_atual + '/' + meses[
            mes_atual].title() + '/' + today
        dia_ontem = int(date.today().strftime('%d')) - 1
        ano_atual_i = int(date.today().strftime('%Y'))
        yest = date(ano_atual_i, int(mes_atual), dia_ontem).strftime("%d-%m-%Y")

        index_stat = ['Número', 'N.Pré-Nota', 'Emissão', 'Fantasia-Destinatário', 'Cidade-Destinatário', 'Uf',
                      'Natureza-Fiscal', 'Situação-Fiscal', 'Descrição-Do-Depósito', 'Fantasia_Do_Transportador',
                      'Fantasia-Comissionado', 'Data-De-Coleta', 'Previsão-Entrega', 'D_Entrega', 'Agendamento',
                      'Lead-Time',
                      'Dias-Para-Entrega', 'Resumo', 'Status']
        path_r = r'K:/CWB/Logistica/Rastreamento/Patrick/Storage/'
        partial_path = r'K:/CWB/Logistica/Rastreamento/Patrick/Storage/' + ano_atual + '/' + meses[
            mes_atual].title() + '/'
        file = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/MONITORAMENTO ' + ano_atual + '.xlsx'
        file_a = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/MONITORAMENTO ' + str(
            int(ano_atual) - 1) + '.xlsx'

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
                     'Fantasia-Comissionado', '1', '2']
        aut_index_a = ['Número', 'N.Pré-Nota', 'Emissão', 'Fantasia-Destinatário', 'Cidade-Destinatário', 'Uf',
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

        itens_ign = ['tracking; ', 'success; ', 'message; ', 'header; ', 'remetente; ', 'destinatario; ', 'items; ',
                     'item; ',
                     'data_hora; ', 'dominio; ', 'filial; ', 'cidade; ', 'ocorrencia; ', 'descricao; ', 'tipo; ',
                     'data_hora_efetiva; ', 'nome_recebedor; ', 'nro_doc_recebedor; ',
                     '?xml version="1.0" encoding="UTF-8" ?; ', '']

        prov_index = ['Empresa', 'PedidoVenda', 'NF', 'DataEmissão', 'Cliente', 'UF',
                      'Cidade', 'N.F._Situação', 'Deposito', 'Transportador', 'Comissionado',
                      'DataColeta', 'DataAgenda', 'DataEntrega', 'ObservaçãoFrete',
                      'Justificativa']
        index_qlik = ['Empresa', 'Pedido de Venda', 'N.F', 'Data Emissão', 'Cliente', 'UF',
                      'Cidade', 'N.F. Situação', 'Deposito', 'Transportador', 'Comissionado',
                      'Data Coleta', 'Data Agenda', 'Data Entrega', 'Observação Frete',
                      'Justificativa']

        full_name_urano = "Urano - " + str(date.today().strftime('%d-%m-%Y')) + ".CSV"

        p_source_path = "C:/Users/patrick.paula/Downloads/"

        source_path_urano = p_source_path + full_name_urano
        dest_path_urano = dest_path + "/" + full_name_urano

        column_veloz = ['Faturamento', 'Data', 'Expedição', 'Número', 'Pre_nota', 'Volumes', 'Transportadora',
                        'NF_Veloz',
                        'Status', 'Data_De_Coleta', 'Observação', 'Atual', 'Pendencia', 'Aguardando_Separação',
                        'Nao_coletados'
                        ]

    @staticmethod
    def columns_to_datetime(self, val):
        val[11] = pd.to_datetime(arg=val[11], errors='coerce', format='%d-%m-%Y')
        print(val[11])
        val[12] = pd.to_datetime(arg=val[12], errors='coerce', format='%d-%m-%Y')
        print(val[12])
        val[13] = pd.to_datetime(arg=val[13], errors='coerce', format='%d-%m-%Y')
        val[14] = pd.to_datetime(arg=val[14], errors='coerce', format='%d-%m-%Y')
        return val[11], val[12], val[13], val[14]

    @staticmethod
    def to_date_time(self, val):
        if len(val) > 1:
            if val[2] == '-':
                val = pd.to_datetime(arg=val, errors='ignore', format='%d-%m-%Y')
            elif val[2] == '/':
                val = pd.to_datetime(arg=val, errors='ignore', format='%d/%m/%Y')
            valor = val

        else:
            valor = '-'
        return valor

    @staticmethod
    def fatiamento(self, val):
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

    @staticmethod
    def data(cls):
        if cls.mes_atual != '01':
            if cls.dia_semana == 0:
                if cls.dia_hoje == 1:
                    dia = cls.dias_meses[int(cls.mes_atual) - 1] - 2
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d/%m/%Y")

                elif cls.dia_hoje == 2:
                    dia = cls.dias_meses[int(cls.mes_atual) - 1] - 1
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d/%m/%Y")

                elif cls.dia_hoje == 3:
                    dia = cls.dias_meses[cls.mes_atual]
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d/%m/%Y")

                else:
                    dia = cls.dia_hoje - 3
                    return date(cls.ano_tual, int(cls.mes_atual), dia).strftime("%d/%m/%Y")

            else:
                if cls.dia_hoje == 1:
                    dia = cls.dias_meses[int(cls.mes_atual) - 1]
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d/%m/%Y")
                else:
                    dia = cls.dia_hoje - 1
                    return date(cls.ano_tual, int(cls.mes_atual), dia).strftime("%d/%m/%Y")
        else:
            if cls.dia_semana == 0:
                if cls.dia_hoje == 1:
                    dia = cls.dias_meses[int(cls.mes_atual) - 1] - 2
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d/%m/%Y")

                elif cls.dia_hoje == 2:
                    dia = cls.dias_meses[int(cls.mes_atual) - 1] - 1
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d/%m/%Y")

                elif cls.dia_hoje == 3:
                    dia = cls.dias_meses[cls.mes_atual]
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d/%m/%Y")

                else:
                    dia = cls.dia_hoje - 3
                    return date(cls.ano_tual, int(cls.mes_atual), dia).strftime("%d/%m/%Y")

            else:
                if cls.dia_hoje == 1:
                    dia = cls.dias_meses[int(cls.mes_atual) - 1]
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d/%m/%Y")
                else:
                    dia = cls.dia_hoje - 1
                    return date(cls.ano_tual, int(cls.mes_atual), dia).strftime("%d/%m/%Y")

    @staticmethod
    def dataf(cls):
        if cls.mes_atual != '01':
            if cls.dia_semana == 0:
                if cls.dia_hoje == 1:
                    dia = cls.dias_meses[int(cls.mes_atual) - 1] - 2
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d-%m-%Y")

                elif cls.dia_hoje == 2:
                    dia = cls.dias_meses[int(cls.mes_atual) - 1] - 1
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d-%m-%Y")

                elif cls.dia_hoje == 3:
                    dia = cls.dias_meses[int(cls.mes_atual)]
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d-%m-%Y")

                else:
                    dia = cls.dia_hoje - 3
                    return date(cls.ano_tual, int(cls.mes_atual), dia).strftime("%d-%m-%Y")
            else:
                if cls.dia_hoje == 1:
                    dia = cls.dias_meses[int(cls.mes_atual) - 1]
                    return date(cls.ano_tual, int(cls.mes_atual) - 1, dia).strftime("%d-%m-%Y")
                else:
                    dia = cls.dia_hoje - 1
                    return date(cls.ano_tual, int(cls.mes_atual), dia).strftime("%d-%m-%Y")
        elif cls.dia_semana == 0:
            if cls.dia_hoje == 1:
                dia = cls.dias_meses[12] - 2
                return date(cls.ano_tual - 1, 12, dia).strftime("%d-%m-%Y")

            elif cls.dia_hoje == 2:
                dia = cls.dias_meses[12] - 2
                return date(cls.ano_tual - 1, 12, dia).strftime("%d-%m-%Y")

            elif cls.dia_hoje == 3:
                dia = cls.dias_meses[12] - 2
                return date(cls.ano_tual - 1, 12, dia).strftime("%d-%m-%Y")
            elif cls.dia_hoje == 4:
                dia = cls.dias_meses[12] - 2
                return date(cls.ano_tual - 1, 12, dia).strftime("%d-%m-%Y")

            else:
                dia = cls.dia_hoje - 3
                return date(cls.ano_tual - 1, 12, dia).strftime("%d-%m-%Y")
        else:
            if cls.dia_hoje == 1:
                dia = cls.dias_meses[12]
                return date(cls.ano_tual - 1, 12, dia).strftime("%d-%m-%Y")
            else:
                dia = cls.dia_hoje - 1
                return date(cls.ano_tual, int(cls.mes_atual), dia).strftime("%d-%m-%Y")

    @staticmethod
    def calc_data(self, dia, mes, ano, valor_a_somar):
        uteis = [0, 1, 2, 3, 4]
        if date.weekday(date(ano, mes, dia)) == 4 and valor_a_somar == 1:
            dia += 3
            valor_a_somar -= 1

        while valor_a_somar > 0:

            data = date.weekday(date(ano, mes, dia))

            if data in uteis:
                if mes == 12 and cls.dias_meses[mes] == dia:
                    dia = 1
                    mes = 1
                    ano += 1
                    valor_a_somar -= 1
                elif self.dias_meses[mes] == dia:
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

    def func(cls, val):
        val_ignore = ['', '-']
        val = val.split('!')
        try:
            if val[0] not in val_ignore and val[1] not in val_ignore:
                val[1] = int(val[1][:-2])
                dia = int(val[0][8:10])
                mes = int(val[0][5:7])
                ano = int(val[0][:4])
                val = cls.calc_data(dia, mes, ano, val[1])
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
                    dia_prev = dias_meses[mes_prev]
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
        print(val)
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

    def dia_ant_ga():
        dia_atual_ga = int(datetime.today().strftime('%d'))
        mes_atual_ga = int(datetime.today().strftime('%m'))
        ano_atual_ga = int(datetime.today().strftime('%Y'))

        dia_semana_ga = datetime.weekday()

        if dia_semana_ga == 0:
            if mes_atual_ga != 1:
                if dia_atual_ga >= 4:
                    if mes_atual_ga < 10:
                        data_1 = '0' + str(dia_atual_ga - 3) + '/0' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                        data_2 = '0' + str(dia_atual_ga - 1) + '/0' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    else:
                        data_1 = '0' + str(dia_atual_ga - 3) + '/' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                        data_2 = '0' + str(dia_atual_ga - 1) + '/' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    ret = [data_1, data_2]
                    return ret
                elif dia_atual_ga == 3:
                    if mes_atual_ga < 10:
                        data_1 = str(dias_meses[mes_atual_ga - 1]) + '/0' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                        data_2 = '0' + str(dia_atual_ga - 1) + '/0' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    elif mes_atual_ga == 10:
                        data_1 = str(dias_meses[mes_atual_ga - 1]) + '/0' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                        data_2 = '0' + str(dia_atual_ga - 1) + '/' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    else:
                        data_1 = str(dias_meses[mes_atual_ga - 1]) + '/' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                        data_2 = '0' + str(dia_atual_ga - 1) + '/' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    ret = [data_1, data_2]
                    return ret
                elif dia_atual_ga == 2:
                    if mes_atual_ga < 10:
                        data_1 = str(dias_meses[mes_atual_ga - 1] - 1) + '/0' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                        data_2 = '0' + str(dia_atual_ga - 1) + '/0' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    elif mes_atual_ga == 10:
                        data_1 = str(dias_meses[mes_atual_ga - 1] - 1) + '/0' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                        data_2 = '0' + str(dia_atual_ga - 1) + '/' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    else:
                        data_1 = str(dias_meses[mes_atual_ga - 1] - 1) + '/' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                        data_2 = '0' + str(dia_atual_ga - 1) + '/' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    ret = [data_1, data_2]
                    return ret
                elif dia_atual_ga == 1:
                    if mes_atual_ga < 10:
                        data_1 = str(dias_meses[mes_atual_ga - 1] - 2) + '/0' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                        data_2 = str(dias_meses[mes_atual_ga - 1]) + '/0' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                    elif mes_atual_ga == 10:
                        data_1 = str(dias_meses[mes_atual_ga - 1] - 2) + '/0' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                        data_2 = str(dias_meses[mes_atual_ga - 1]) + '/0' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                    else:
                        data_1 = str(dias_meses[mes_atual_ga - 1] - 2) + '/' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                        data_2 = str(dias_meses[mes_atual_ga - 1]) + '/' + str(mes_atual_ga - 1) + '/' + str(
                            ano_atual_ga)
                    ret = [data_1, data_2]
                    return ret
            else:  # Se o mês for Janeiro
                if dia_atual_ga >= 4:
                    data_1 = '0' + str(dia_atual_ga - 3) + '/0' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    data_2 = '0' + str(dia_atual_ga - 1) + '/0' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    ret = [data_1, data_2]
                    return ret
                elif dia_atual_ga == 3:
                    data_1 = str(dias_meses[12]) + '/' + str(12) + '/' + str(ano_atual_ga - 1)
                    data_2 = '0' + str(dia_atual_ga - 1) + '/0' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    ret = [data_1, data_2]
                    return ret
                elif dia_atual_ga == 2:
                    data_1 = str(dias_meses[12] - 1) + '/' + str(12) + '/' + str(ano_atual_ga - 1)
                    data_2 = '0' + str(dia_atual_ga - 1) + '/0' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                    ret = [data_1, data_2]
                    return ret
                elif dia_atual_ga == 1:
                    data_1 = str(dias_meses[12] - 2) + '/' + str(12) + '/' + str(ano_atual_ga - 1)
                    data_2 = str(dias_meses[12]) + '/' + str(12) + '/' + str(ano_atual_ga - 1)
                    ret = [data_1, data_2]
                    return ret
        else:
            if mes_atual_ga < 10:
                if dia_atual_ga < 10:
                    data_1 = '0' + str(dia_atual_ga - 1) + '/0' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                else:
                    data_1 = str(dia_atual_ga - 1) + '/0' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                ret = [data_1, data_1]
                return ret
            else:
                if dia_atual_ga < 10:
                    data_1 = '0' + str(dia_atual_ga - 1) + '/' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                else:
                    data_1 = str(dia_atual_ga - 1) + '/' + str(mes_atual_ga) + '/' + str(ano_atual_ga)
                ret = [data_1, data_1]
                return ret

    ac_mon = date.today().strftime("%m")
    ac_yea = date.today().strftime("%Y")
    ac_day = int(dia_atual)
    ac_day = str(ac_day)
    partial_name_1 = "Qlik Sense - Planilha de Alimentação de Datas - "
    partial_name_2 = ".xlsx"
    de = " de "
    full_name_qlik = partial_name_1 + ac_day + de + meses[ac_mon] + de + ac_yea + partial_name_2
    today = date.today().strftime("%d-%m-%Y")
    p_source_path = "C:/Users/patrick.paula/Downloads/"

    source_path_qlik = p_source_path + full_name_qlik
    dest_path_qlik = dest_path + "/" + full_name_qlik
