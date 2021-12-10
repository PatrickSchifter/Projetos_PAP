import pywhatkit
import pandas as pd
import os

dfs = []
dir_path = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Notas/'
path_min = r'K:/CWB/Logistica/Rastreamento/Minutas_coleta/'
cli_path = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Clientes.xls'
df_com = r'K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Representantes.xlsx'

for item in list(os.listdir(dir_path)):
    if 'Clientes' in item:
        pass
    else:
        df = pd.read_excel(dir_path + item)
        dfs.append(df)

df_not = pd.read_excel(dir_path[:-6] + 'Notas/Notas_Clientes.xls')

dfs.append(df_not)

for x in range(0, len(dfs)):
    df_not = pd.concat(objs=dfs, join='outer')

df_min = pd.DataFrame
for item in list(os.listdir(path_min)):
    df_min = pd.read_excel(path_min + item)
df_cli = pd.read_excel(cli_path)
df_com = pd.read_excel(df_com)

df_min.columns = ['Data', 'Número', 'Volumes', 'Transportador']
df_min = df_min[['Data', 'Número']]
df_min = df_min[4:]
frame = pd.merge(left=df_min, right=df_not, how='left', on='Número')
frame = frame.fillna('0')
frame = frame.query("Destinatário != '0'")
df_cli.columns = ['Destinatário', 'Fantasia', 'Razão Social', 'CNPJ/CPF', 'Cidade', 'UF', 'Fone', 'Celular Cliente',
                  'E-mail Corporativo', 'Situação', 'Tipo Cliente', 'Grupo Econômico', 'Nota']
frame = pd.merge(left=frame, right=df_cli, how='left', on='Destinatário')
frame = pd.merge(left=frame, right=df_com, how='left', on='Fantasia Comissionado')
frame = frame[['Data', 'Número', 'N. Pré Nota', 'Fantasia Destinatário', 'Fantasia Comissionado', 'Razão Social',
               'Celular Cliente', 'Celular']]
frame.columns = ['Data', 'Número', 'N_Pre_Nota', 'Fantasia Destinatário', 'Fantasia_Comissionado', 'Razao_Social',
                 'Cel_Cliente', 'Celular_Comissionado']
frame = frame.astype({'N_Pre_Nota': 'int64'})
frame = frame.fillna(0)
lista_envios = ['BRUNA TEBALDI', 'PAULO HENRIQUE', 'ARYADNE RONCAGLIO MARTINS']
print(frame)

meu_num = '+5541991912238'
for row in frame.itertuples(index=True, name='Row'):
    if row.Fantasia_Comissionado in lista_envios:
        tel_cli = '+' + str(row.Cel_Cliente)
        msg_cli = 'Prezado cliente, informamos que seu pedido com a Porto a Porto NF {} já saiu para entrega e em breve estará chegando em seu estabelecimento! \n\n\nObs: Essa é uma mensagem automática, por gentileza não responder.'.format(
            row.Número)
        pywhatkit.sendwhatmsg_instantly(phone_no=tel_cli,
                                        message=msg_cli, wait_time=5,
                                        tab_close=True, close_time=2)

        tel_com = '+' + str(row.Celular_Comissionado)
        msg_com = "Prezado representante, informamos que o pedido do cliente {}, NF {} (Pré Nota - {}), está em rota e em breve será entregue! \n\n\nObs: Essa é uma mensagem automática, por gentileza não responder.".format(
            row.Razao_Social, row.Número, row.N_Pre_Nota)
        pywhatkit.sendwhatmsg_instantly(phone_no=tel_com, wait_time=5,
                                        message=msg_com,
                                        tab_close=True, close_time=2)

pywhatkit.sendwhatmsg_instantly(phone_no=meu_num, message='Mensagens enviadas', wait_time=5, tab_close=True,
                                close_time=2)

for item in list(os.listdir(path_min)):
    os.remove(path_min + item)
