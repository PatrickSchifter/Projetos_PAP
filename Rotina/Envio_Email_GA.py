import win32com.client as win32
from config import ano_atual

outlook = win32.Dispatch('outlook.application')
file_ga = 'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/MONITORAMENTO GA ' + ano_atual + '.xlsx'
mail = outlook.CreateItem(0)
mail.To = 'melissa.crocetti@portoaporto.com.br; aline.gomes@portoaporto.com.br; andre.pereira@grandeadega.com.br; joao.oliveira@grandeadega.com.br; rafaelli.nascimento@grandeadega.com.br; luiz.souza@grandeadega.com.br; luana.neves@portoaporto.com.br; camila.podolak@portoaporto.com.br; camila@grandeadega.com.br; rafael.felczak@portoaporto.com.br'
mail.Subject = 'Controle de monitoramento Grande Adega ' + ano_atual
mail.Body = 'Prezados, bom dia!!! \n\nSegue Controle de Monitoramento Grande Adega.\n\n Esse é um email automático'
mail.Attachments.Add(file_ga)
mail.Send()
