import pandas as pd

ceps = 'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Aquivos Grande Adega/ceps.txt'
jadlog = 'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Aquivos Grande Adega/Jadlog.txt'
PAC = 'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Aquivos Grande Adega/PAC.txt'
sedex = 'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Aquivos Grande Adega/SEDEX.txt'
tnt = 'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Aquivos Grande Adega/TNT.txt'
final_file = 'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Aquivos Grande Adega/Prazos.txt'

bd_ceps = pd.read_csv(filepath_or_buffer=ceps, sep='	', header=None, error_bad_lines=True)
bd_jadlog = pd.read_csv(filepath_or_buffer=jadlog, sep='	', header=None, error_bad_lines=True)
bd_PAC = pd.read_csv(filepath_or_buffer=PAC, sep='	', header=None, error_bad_lines=True)
bd_sedex = pd.read_csv(filepath_or_buffer=sedex, sep='	', header=None, error_bad_lines=True)
bd_tnt = pd.read_csv(filepath_or_buffer=tnt, sep='	', header=None, error_bad_lines=True)
bd_ceps.columns = ['CEP', 'Cidade/UF', 'Bairro']
bd_ceps['CEP'] = bd_ceps['CEP'].apply(lambda val: pd.to_numeric(val))

bd_jadlog.columns = ['CEP_inicio', 'CEP_fim', 'PolygonName', 'WeightStart', 'WeightEnd', 'AbsoluteMoneyCost',
                     'PricePercent', 'PriceByExtraWeight', 'MaxVolume', 'Lead_Time', 'Country',
                     'MinimumValueInsurance']
bd_jadlog = bd_jadlog[['CEP_inicio', 'CEP_fim', 'Lead_Time']]
bd_jadlog[['CEP_inicio', 'CEP_fim']] = bd_jadlog[['CEP_inicio', 'CEP_fim']].apply(
    func=lambda val: pd.to_numeric(val))
bd_jadlog['Lead_Time'] = bd_jadlog['Lead_Time'].apply(lambda val: val.split('.'))
bd_jadlog['Lead_Time'] = bd_jadlog['Lead_Time'].apply(lambda val: val[0])
bd_jadlog['Lead_Time'] = bd_jadlog['Lead_Time'].apply(lambda val: pd.to_numeric(val))
bd_jadlog = bd_jadlog.drop_duplicates(subset=['CEP_inicio'])

bd_PAC.columns = ['CEP_inicio', 'CEP_fim', 'PolygonName', 'WeightStart', 'WeightEnd', 'AbsoluteMoneyCost',
                  'PricePercent', 'PriceByExtraWeight', 'MaxVolume', 'Lead_Time', 'Country',
                  'MinimumValueInsurance']
bd_PAC = bd_PAC[['CEP_inicio', 'CEP_fim', 'Lead_Time']]
bd_PAC = bd_PAC.drop_duplicates(subset=['CEP_inicio'])
bd_PAC[['CEP_inicio', 'CEP_fim']] = bd_PAC[['CEP_inicio', 'CEP_fim']].apply(
    func=lambda val: pd.to_numeric(val))
bd_PAC['Lead_Time'] = bd_PAC['Lead_Time'].apply(lambda val: val.split('.'))
bd_PAC['Lead_Time'] = bd_PAC['Lead_Time'].apply(lambda val: val[0])
bd_PAC['Lead_Time'] = bd_PAC['Lead_Time'].apply(lambda val: pd.to_numeric(val))

bd_sedex.columns = ['CEP_inicio', 'CEP_fim', 'PolygonName', 'WeightStart', 'WeightEnd', 'AbsoluteMoneyCost',
                    'PricePercent', 'PriceByExtraWeight', 'MaxVolume', 'Lead_Time', 'Country',
                    'MinimumValueInsurance']
bd_sedex = bd_sedex[['CEP_inicio', 'CEP_fim', 'Lead_Time']]
bd_sedex = bd_sedex.drop_duplicates(subset=['CEP_inicio'])
bd_sedex[['CEP_inicio', 'CEP_fim']] = bd_sedex[['CEP_inicio', 'CEP_fim']].apply(
    func=lambda val: pd.to_numeric(val))
bd_sedex['Lead_Time'] = bd_sedex['Lead_Time'].apply(lambda val: val.split('.'))
bd_sedex['Lead_Time'] = bd_sedex['Lead_Time'].apply(lambda val: val[0])
bd_sedex['Lead_Time'] = bd_sedex['Lead_Time'].apply(lambda val: pd.to_numeric(val))

bd_tnt.columns = ['CEP_inicio', 'CEP_fim', 'PolygonName', 'WeightStart', 'WeightEnd', 'AbsoluteMoneyCost',
                  'PricePercent', 'PriceByExtraWeight', 'MaxVolume', 'Lead_Time', 'Country',
                  'MinimumValueInsurance']
bd_tnt = bd_tnt[['CEP_inicio', 'CEP_fim', 'Lead_Time']]
bd_tnt = bd_tnt.drop_duplicates(subset=['CEP_inicio'])
bd_tnt[['CEP_inicio', 'CEP_fim']] = bd_tnt[['CEP_inicio', 'CEP_fim']].apply(
    func=lambda val: pd.to_numeric(val))
bd_tnt['Lead_Time'] = bd_tnt['Lead_Time'].apply(lambda val: val.split('.'))
bd_tnt['Lead_Time'] = bd_tnt['Lead_Time'].apply(lambda val: val[0])
bd_tnt['Lead_Time'] = bd_tnt['Lead_Time'].apply(lambda val: pd.to_numeric(val))


def localizador_leadtime(val, id_transp):
    valor = '-'
    if id_transp == 1:
        for row in bd_jadlog.itertuples(index=True, name='Row'):
            if val > row.CEP_inicio and val < row.CEP_fim:
                print(val, ' - Cep inicio: ', row.CEP_inicio, ' - Cep fim : ', row.CEP_fim, ' - Lead-Time: ', row.Lead_Time)
                try:
                    valor = row.Lead_Time
                    if valor > 0:
                        continue
                except UnboundLocalError:
                    valor = '-'
    if id_transp == 2:
        for row in bd_PAC.itertuples(index=True, name='Row'):
            if val > row.CEP_inicio and val < row.CEP_fim:
                print(val, ' - Cep inicio: ', row.CEP_inicio, ' - Cep fim : ', row.CEP_fim, ' - Lead-Time: ', row.Lead_Time)
                try:
                    valor = row.Lead_Time
                    if valor > 0:
                        continue
                except UnboundLocalError:
                    valor = '-'
    if id_transp == 3:
        for row in bd_sedex.itertuples(index=True, name='Row'):
            if val > row.CEP_inicio and val < row.CEP_fim:
                print(val, ' - Cep inicio: ', row.CEP_inicio, ' - Cep fim : ', row.CEP_fim, ' - Lead-Time: ', row.Lead_Time)
                try:
                    valor = row.Lead_Time
                    if valor > 0:
                        continue
                except UnboundLocalError:
                    valor = '-'
    if id_transp == 4:
        for row in bd_tnt.itertuples(index=True, name='Row'):
            if val > row.CEP_inicio and val < row.CEP_fim:
                print(val, ' - Cep inicio: ', row.CEP_inicio, ' - Cep fim : ', row.CEP_fim, ' - Lead-Time: ', row.Lead_Time)
                try:
                    valor = row.Lead_Time
                    if valor > 0:
                        continue
                except UnboundLocalError:
                    valor = '-'
    return valor

if __name__ == '__main__' :
    bd_ceps['Lead_time_Jadlog'] = bd_ceps['CEP'].apply(func=lambda val: localizador_leadtime(val, 1))
    bd_ceps['Lead_time_PAC'] = bd_ceps['CEP'].apply(func=lambda val: localizador_leadtime(val, 2))
    bd_ceps['Lead_time_Sedex'] = bd_ceps['CEP'].apply(func=lambda val: localizador_leadtime(val, 3))
    bd_ceps['Lead_time_TNT'] = bd_ceps['CEP'].apply(func=lambda val: localizador_leadtime(val, 4))

    bd_ceps.to_csv(path_or_buf=final_file, index=False)

    for x in bd_ceps.iterrows():
        print(x)