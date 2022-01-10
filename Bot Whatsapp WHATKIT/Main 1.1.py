import pywhatkit
import pandas as pd
import os

dfs = []
dir_path = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Notas/'
path_min = r'C:\Users\patrick.paula\Porto a Porto Comercio de IMP e EXP LTDA\Afonso Marcon - Minutas_Coleta/'
cli_path = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Clientes.xls'
df_rep = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Representantes.xlsx'

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

for item in list(os.listdir(path_min)):
    print(item)
    try:
        df_min = pd.read_excel(path_min + item)
    except PermissionError:
        continue

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
df_cli.columns = ['Destinatário', 'Fantasia', 'Razão Social', 'CNPJ/CPF', 'Cidade', 'UF', 'Fone', 'Celular Cliente',
                  'E-mail Corporativo', 'Situação', 'Tipo Cliente', 'Grupo Econômico', 'Nota']
frame = pd.merge(left=frame, right=df_cli, how='left', on='Destinatário')
frame = pd.merge(left=frame, right=df_rep, how='left', on='Fantasia Comissionado')
frame = frame[['Número', 'N. Pré Nota', 'Fantasia Destinatário', 'Fantasia Comissionado', 'Razão Social',
               'Celular Cliente', 'Celular']]
frame.columns = ['Número', 'N_Pre_Nota', 'Fantasia Destinatário', 'Fantasia_Comissionado', 'Razao_Social',
                 'Cel_Cliente', 'Celular_Comissionado']

frame = frame.astype({'N_Pre_Nota': 'int64'})
frame = frame.fillna(0)
lista_envios = ['BRUNA TEBALDI', 'PAULO HENRIQUE', 'ARYADNE RONCAGLIO MARTINS']
#
print(frame['Fantasia_Comissionado'])

meu_num = '+5541991912238'
frame = frame.drop_duplicates('Número','first')
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

for item in list(os.listdir(path_min)):
    try:

        os.remove(path_min + item)
    except PermissionError:
        continue
