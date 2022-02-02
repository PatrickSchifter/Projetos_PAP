import requests
import pandas as pd
from Projetos.Rotina.config import itens_ign, all_index, indexes, meses_str, mes_atual, ano_atual, dia_atual
import time
start_time = time.time()

url = 'https://ssw.inf.br/api/tracking'
cnpj = '00069957000194'
file_m_atual = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/MONITORAMENTO ' + ano_atual + '.xlsx'
file_m_ant = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/MONITORAMENTO' + str(int(ano_atual) - 1) + 'xlsx'
if dia_atual == '01':
    name_m_atual = meses_str[str(int(mes_atual) - 1)].upper() + '-' + str(ano_atual)
else:
    name_m_atual = meses_str[str(int(mes_atual))].upper() + '-' + str(ano_atual)
if mes_atual == '1':
    name_m_ant = meses_str['12'].upper() + '-' + str(int(ano_atual) - 1)
else:
    try:
        name_m_ant = meses_str[str(int(mes_atual) - 1)].upper() + '-' + str(ano_atual)
    except KeyError:
        # Já foi tratada a exceção no if
        pass


def search_data(mes):
    df = pd.read_excel(file_m_atual, mes)
    df.columns = all_index
    df.fillna('-')
    df = df.query(
        "Fantasia_Do_Transportador != 'MALBEC' & Fantasia_Do_Transportador != 'URANOLOG' &\
        Fantasia_Do_Transportador != 'JOHNKE TRANSPORTES' & Fantasia_Do_Transportador != '-' &\
         Fantasia_Do_Transportador != 'ALIANCA' & Fantasia_Do_Transportador != 'GRANDE ADEGA - MATRIZ' &\
          Fantasia_Do_Transportador != 'TRANSFRIOS TRANSP' & Fantasia_Do_Transportador != 'BRINGER DO BRASIL' &\
           Fantasia_Do_Transportador != 'GRANDE ADEGA' & Fantasia_Do_Transportador != 'TRANSFRIOS SP' & D_Entrega == '-'")
    nf = []
    conta = 0
    for x in df['Número']:
        if conta == 20:
            time.sleep(1 / 3)
            conta = 0
        conta += 1

        nro_nf = x

        request = requests.post(url=url, data={'cnpj': cnpj, 'nro_nf': nro_nf})

        info = []
        inf = []
        cont = 0
        for stat in request.text.split(sep='<'):

            stat = stat.replace('/', '')
            stat = stat.replace('>', '; ')
            if 'success; false' in stat:
                continue
            if stat in itens_ign:
                continue
            elif 'tipo' in stat:
                continue
            elif 'tracking' in stat:
                continue
            elif 'message' in stat:
                continue
            elif 'remetente' in stat:
                continue
            elif 'destinatario' in stat:
                continue
            elif 'success' in stat:
                continue
            elif 'efetiva' in stat:
                continue

            stat = stat.split(';')
            if cont == 0:
                if stat[1][1].isnumeric():
                    inf.append(nro_nf)
                    stat[1] = stat[1][1:11]
                else:
                    continue

            if cont == 5:
                cont = 0
                inf.append(stat[1])
                info.append(inf)
                inf = []
            else:
                inf.append(stat[1])
                cont += 1

        try:
            nf.append(info[-1])
        except IndexError:
            pass
            info = []

    df_ssw = pd.DataFrame(data=nf, columns=indexes)
    df_ssw['Data'] = pd.to_datetime(df_ssw['Data'])

    return df_ssw


print("--- %s seconds SSW integration ---" % (time.time() - start_time))
