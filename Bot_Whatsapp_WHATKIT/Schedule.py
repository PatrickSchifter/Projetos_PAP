from datetime import datetime
import time
from Down_Notas_Emitidas import down_notas
import os
from Exec import enviar_msg

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
               '14:11', '14:12', '14:13', '14:14', '14:15',
               '15:00', '15:01', '15:02', '15:03', '15:04', '15:05', '15:06', '15:07', '15:08', '15:09', '15:10',
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
            enviar_msg()
        except:
            print('Não há nada para enviar')
            time.sleep(60)

    elif hour in horario_down and day in dias_uteis:
        down_notas()

    time.sleep(15)
