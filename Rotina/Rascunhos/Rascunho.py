import win32com.client as win32
from datetime import datetime
import os

outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
data_hoje = datetime.today().strftime('%d/%m/%Y')
subject = 'CONTROLE DE PEDIDOS PAP 01/02/2022'

for message in inbox.Items:
    if message.Subject == subject:
        for att in message.Attachments:
            att.SaveAsFile(os.path.join('C:/Users/patrick.paula/Desktop', att.FileName))
