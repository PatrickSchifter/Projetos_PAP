from datetime import date
from datetime import datetime

dia_hoje = int(datetime.today().strftime('%d'))
mes_hoje = int(datetime.today().strftime('%m'))
ano_hoje = int(datetime.today().strftime('%Y'))


print(type(dia_hoje))