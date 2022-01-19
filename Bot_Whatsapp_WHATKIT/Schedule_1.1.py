from datetime import datetime
import time
from Exec import enviar_msg
from Down_Notas_Emitidas import down_notas

hora_inicio = ["08:00", "08:01", "08:02", "08:03", "08:04"]
dias_uteis = [dias for dias in range(0, 6)]

while True:
    today = datetime.weekday(datetime.today())
    if today in dias_uteis:
        time_now = datetime.now().strftime('%H:%M')
        if time in hora_inicio:
            for hour in range(8, 18):
                times = 0
                for min in range(0, 16):
                    times += 1
                    enviar_msg()
                    time.sleep(60)
                    if times == 16:
                        if hour == 12:
                            down_notas()
                        sec_to_hour = (60 - int(datetime.now().strftime('%M'))) * 60
                        time.sleep(sec_to_hour)
            for night_hour in range(0, 13):
                if night_hour == 4:
                    down_notas()
                    sec_to_hour = (60 - int(datetime.now().strftime('%M'))) * 60
                    time.sleep(sec_to_hour)
                    continue
                time.sleep(3600)
        else:
            time.sleep(30)
    else:
        time.sleep(3600)

