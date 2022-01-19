import win32com.client as win32
from datetime import datetime

def regra_cump():
    msg = ''
    hour = int(datetime.today().strftime('%H'))
    if hour in range(0, 13):
        msg = 'Prezados, bom dia!!!\nForam enviados pela automação de WhatsApp mensagem para os seguintes clientes:\n'
    elif hour in range(13, 18):
        msg = 'Prezados, boa tarde!!!\nForam enviados pela automação de WhatsApp mensagem para os seguintes clientes:\n'
    elif hour in range(18, 00):
        msg = 'Prezados, boa noite!!!\nForam enviados pela automação de WhatsApp mensagem para os seguintes clientes:\n'

    return msg



def enviar_email(destinatario, body_email):
    today = datetime.today().strftime('%d/%m/%Y - %H:%M')
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = destinatario
    mail.Subject = 'Envio de mensagens Automação Whatsapp ' + today
    mail.Body = regra_cump() + '\n' + body_email + '\n\nEssa é uma mensagem automática'
    mail.Send()
