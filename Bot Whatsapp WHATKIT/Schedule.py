from datetime import datetime
import time
import pywhatkit
from Down_Notas_Emitidas import down_notas
import os
import pandas as pd

path_min = r'C:\Users\patrick.paula\Porto a Porto Comercio de IMP e EXP LTDA\Afonso Marcon - Minutas_Coleta/'
dir_path = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Notas/'
cli_path = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Clientes.xls'
df_rep = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Representantes.xlsx'

horario_com = ['08:00', '08:01', '08:02', '08:03', '08:04', '08:05', '08:06', '08:07', '08:08', '08:09', '08:10',
               '08:11', '08:12', '08:13', '08:14', '08:15',
               '09:00', '09:01', '09:02', '09:03', '09:04', '09:05', '09:06', '09:07', '09:08', '09:09', '09:10',
               '09:11', '09:12', '09:13', '09:14', '09:15',
               '10:00', '10:01', '10:02', '10:03', '10:04', '10:05', '10:06', '10:07', '10:08', '10:09', '10:10',
               '10:11', '10:12', '10:13', '10:14', '10:15',
               '11:00', '11:01', '11:02', '11:03', '11:04', '11:05', '11:06', '11:07', '11:08', '11:09', '11:10',
               '11:11', '11:12', '11:13', '11:14', '11:15',
               '13:00', '13:01', '13:02', '13:03', '13:04', '13:05', '13:06', '13:07', '13:08', '13:09', '13:10',
               '13:11', '13:12', '13:13', '13:14', '13:15',
               '14:00', '14:01', '14:02', '14:03', '14:04', '14:05', '14:06', '14:07', '14:08', '14:09', '14:10',
               '14:11', '14:12', '14:13', '14:14', '14:15', '14:25'
                                                            '15:00', '15:01', '15:02', '15:03', '15:04', '15:05',
               '15:06', '15:07', '15:08', '15:09', '15:10',
               '15:11', '15:12', '15:13', '15:14', '15:15',
               '16:00', '16:01', '16:02', '16:03', '16:04', '16:05', '16:06', '16:07', '16:08', '16:09', '16:10',
               '16:11', '16:12', '16:13', '16:14', '16:15',
               '17:00', '17:01', '17:02', '17:03', '17:04', '17:05', '17:06', '17:07', '17:08', '17:09', '17:10',
               '17:11', '17:12', '17:13', '17:14', '17:15',
               '18:00', '18:01', '18:02', '18:03', '18:04', '18:05', '18:06', '18:07', '18:08', '18:09', '18:10',
               '18:11', '18:12', '18:13', '18:14', '18:15']
dias_uteis = [0, 1, 2, 3, 4, 5]

horario_down = ['23:15', '12:15']
while True:
    list_min = list(os.listdir(path_min))
    hour = datetime.now().strftime('%H:%M')
    print(hour)
    day = datetime.weekday(datetime.today())
    if hour in horario_com and day in dias_uteis:
        try:
            for item in list_min:
                print(item)
                try:
                    df_min = pd.read_excel(path_min + item)
                    time.sleep(3)
                    os.remove(path_min + item)
                except PermissionError:
                    continue
            dfs = []

            for item in list(os.listdir(dir_path)):
                if 'Clientes' in item:
                    pass
                else:
                    df = pd.read_excel(dir_path + item)
                    dfs.append(df)
            try:
                df_not = pd.read_excel(dir_path[:-6] + 'Notas/Notas_Clientes.xls')
                dfs.append(df_not)
            except:
                pass

            for x in range(0, len(dfs)):
                df_not = pd.concat(objs=dfs, join='outer')

            df_min = pd.DataFrame

            df_cli = pd.read_excel(cli_path)

            df_rep = pd.read_excel(df_rep)

            wait_time = 20
            print(df_min)

            try:
                df_min.columns = ['Número', 'Data']
            except ValueError:
                try:
                    df_min.columns = ['Número', 'Volumes', 'Transportador']
                except ValueError:
                    df_min.columns = ['Número', 'Volumes', 'Transportador', 'Data']

            df_min = df_min['Número']
            frame = pd.merge(left=df_min, right=df_not, how='left', on='Número')
            frame = frame.fillna('0')
            frame = frame.query("Destinatário != '0'")
            df_cli.columns = ['Destinatário', 'Fantasia', 'Razão Social', 'CNPJ/CPF', 'Cidade', 'UF', 'Fone',
                              'Celular Cliente',
                              'E-mail Corporativo', 'Situação', 'Tipo Cliente', 'Grupo Econômico', 'Nota']
            frame = pd.merge(left=frame, right=df_cli, how='left', on='Destinatário')
            frame = pd.merge(left=frame, right=df_rep, how='left', on='Fantasia Comissionado')
            frame = frame[
                ['Número', 'N. Pré Nota', 'Fantasia Destinatário', 'Fantasia Comissionado', 'Razão Social',
                 'Celular Cliente', 'Celular']]
            frame.columns = ['Número', 'N_Pre_Nota', 'Fantasia Destinatário', 'Fantasia_Comissionado',
                             'Razao_Social',
                             'Cel_Cliente', 'Celular_Comissionado']

            frame = frame.astype({'N_Pre_Nota': 'int64'})
            frame = frame.fillna(0)
            lista_envios = ['BRUNA TEBALDI', 'PAULO HENRIQUE', 'ARYADNE RONCAGLIO MARTINS']
            #
            print(frame['Fantasia_Comissionado'])

            meu_num = '+5541991912238'
            frame = frame.drop_duplicates('Número', 'first')
            print(frame)

            for row in frame.itertuples(index=True, name='Row'):
                tel_cli = '+' + str(row.Cel_Cliente)
                msg_cli = 'Prezado cliente, informamos que seu pedido com a Porto a Porto NF {} já saiu para entrega e em breve estará chegando em seu estabelecimento! \n\n\nObs: Essa é uma mensagem automática, por gentileza não responder.'.format(
                    row.Número)
                pywhatkit.sendwhatmsg_instantly(phone_no=tel_cli,
                                                message=msg_cli, wait_time=wait_time,
                                                tab_close=True, close_time=2)

            msg_vz = 'Prezado representante, informamos que as seguintes notas estão em rota de entrega: \n \nEssa é uma mensagem automática, por gentileza não responder.'

            for rep in df_rep['Fantasia Comissionado']:
                n_frame = frame.query(f"Fantasia_Comissionado == '{str(rep)}'")
                msg = 'Prezado representante, informamos que as seguintes notas estão em rota de entrega: \n \n'

                for row in n_frame.itertuples(index=True, name='Row'):
                    msg += f'NF {row.Número} / Pré-Nota {row.N_Pre_Nota} - {row.Razao_Social} \n \n'

                msg += 'Essa é uma mensagem automática, por gentileza não responder.'

                if msg != msg_vz:
                    print(row.Celular_Comissionado)
                    tel_com = '+' + str(row.Celular_Comissionado)
                    pywhatkit.sendwhatmsg_instantly(phone_no=tel_com, wait_time=wait_time,
                                                    message=msg,
                                                    tab_close=True, close_time=2)
                    msg = ''
                else:
                    msg = ''
                    continue
        except:
            print('Não há nada para enviar')
            time.sleep(120)

    elif hour in horario_down and day in dias_uteis:
        down_notas()

    time.sleep(15)
