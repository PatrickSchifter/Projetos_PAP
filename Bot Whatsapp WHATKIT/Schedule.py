from datetime import datetime
import time

horario_com = ['08:00', '08:15', '09:00', '09:15', '10:00', '10:15', '11:00', '11:15', '12:00', '12:15', '13:00',
               '13:15', '14:00', '14:15', '15:00', '15:15', '16:00', '16:15', '17:00', '17:15', '18:00', '18:15']
dias_uteis = [0, 1, 2, 3, 4, 5]

horario_down = ['23:15', '12:15']


while True:
    hour = datetime.now().strftime('%H:%M')
    day = datetime.weekday(datetime.today())
    time.sleep(15)
    if hour in horario_com and day in dias_uteis:
        for x in range(5):
            try:
                import Executável
            except TypeError:
                pass

    elif hour in horario_down and day in dias_uteis:
        import Down_Notas_Emitidas



