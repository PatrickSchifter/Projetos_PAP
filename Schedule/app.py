import schedule
import time

def teste():
    return print('Esse é um testes')

schedule.every(10).seconds.do(teste())

while True:
    schedule.run_pending()
    time.sleep(1)