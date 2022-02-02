import time
import pywhatkit
import pandas as pd
import os
from Projetos.Bot_Whatsapp_WHATKIT.Enviar_email import enviar_email


def enviar_msg():
    dir_path = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Notas/'
    path_min = r'C:/Users/patrick.paula/Porto a Porto Comercio de IMP e EXP LTDA\Afonso Marcon - Minutas_Coleta/'
    cli_path = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Clientes.xls'
    df_rep = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Representantes.xlsx'
    df_resp = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Responsaveis.xlsx'
    df_resp = pd.read_excel(io=df_resp, index_col=0)

    list_min = list(os.listdir(path_min))
    for item in list_min:
        try:
            df_min = pd.read_excel(path_min + item)
            time.sleep(3)
            os.remove(path_min + item)
        except PermissionError:
            continue

    dfs = []
    for item in list(os.listdir(dir_path)):
        df = pd.read_excel(dir_path + item)
        dfs.append(df)

    for x in range(0, len(dfs)):
        df_not = pd.concat(objs=dfs, join='outer')

    df_cli = pd.read_excel(cli_path)
    df_rep = pd.read_excel(df_rep)

    wait_time = 20

    try:
        try:
            df_min.columns = ['Número', 'Data']
        except ValueError:
            try:
                df_min.columns = ['Número', 'Volumes', 'Transportador']
            except ValueError:
                df_min.columns = ['Número', 'Volumes', 'Transportador', 'Data']

        df_min = df_min['Número']

        frame = pd.merge(left=df_min, right=df_not, how='left', on='Número')
        print(frame)
        frame = frame.fillna('0')
        frame = frame.query("Destinatário != '0'")
        df_cli.columns = ['Destinatário', 'Fantasia', 'Razão Social', 'CNPJ/CPF', 'Cidade', 'UF', 'Fone',
                          'Celular Cliente',
                          'E-mail Corporativo', 'Situação', 'Tipo Cliente', 'Grupo Econômico', 'Nota']
        frame = pd.merge(left=frame, right=df_cli, how='left', on='Destinatário')
        frame = pd.merge(left=frame, right=df_rep, how='left', on='Fantasia Comissionado')
        frame = frame[['Número', 'N. Pré Nota', 'Fantasia Destinatário', 'Fantasia Comissionado', 'Razão Social',
                       'Celular Cliente', 'Celular']]
        frame.columns = ['Número', 'N_Pre_Nota', 'Fantasia Destinatário', 'Fantasia_Comissionado', 'Razao_Social',
                         'Cel_Cliente', 'Celular_Comissionado']

        frame = frame.astype({'N_Pre_Nota': 'int64'})
        frame = frame.fillna(0)

        meu_num = '+5541991912238'
        frame = frame.drop_duplicates('Número', 'first')

        body_email = ''

        print(frame)

        for row in frame.itertuples(index=True, name='Row'):
            tel_cli = '+' + str(row.Cel_Cliente)
            msg_cli = 'Prezado cliente, informamos que seu pedido com a Porto a Porto NF {} já saiu para entrega e em breve estará chegando em seu estabelecimento! \n\n\nObs: Essa é uma mensagem automática, por gentileza não responder.'.format(
                row.Número)
            pywhatkit.sendwhatmsg_instantly(phone_no=tel_cli,
                                            message=msg_cli, wait_time=wait_time,
                                            tab_close=True, close_time=2)
            body_email += f'NF {row.Número} - Pré-Nota {row.N_Pre_Nota} - {row.Razao_Social} \n'

        msg_vz = 'Prezado representante, informamos que as seguintes notas estão em rota de entrega: \n \nEssa é uma mensagem automática, por gentileza não responder.'

        for rep in df_rep['Fantasia Comissionado']:
            n_frame = frame.query(f"Fantasia_Comissionado == '{str(rep)}'")
            msg = 'Prezado representante, informamos que as seguintes notas estão em rota de entrega: \n \n'

            for row in n_frame.itertuples(index=True, name='Row'):
                msg += f'NF {row.Número} / Pré-Nota {row.N_Pre_Nota} - {row.Razao_Social} \n \n'

            msg += 'Essa é uma mensagem automática, por gentileza não responder.'

            if msg != msg_vz:
                print(row.Fantasia_Comissionado)
                representante = row.Fantasia_Comissionado
                tel_com = '+' + str(row.Celular_Comissionado)
                print(tel_com)
                pywhatkit.sendwhatmsg_instantly(phone_no=tel_com, wait_time=wait_time,
                                                message=msg,
                                                tab_close=True, close_time=2)
                msg = ''
            else:
                msg = ''
                continue
    except UnboundLocalError:
        pass
    enviar_email(df_resp.loc[representante]['email'], body_email)
