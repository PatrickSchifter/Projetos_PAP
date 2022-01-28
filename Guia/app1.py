

def enviar_guia():
    from down_everest import down_notas_e, nome_saldos
    import pandas as pd
    from config import capacity_file, guia_file, format_float, regra_cump
    down_notas_e()
    df_cap = pd.read_excel(io=capacity_file, sheet_name='Capacidade')
    df_cap.columns = ['Item', 'Descrição', 'Capacidade_Caixas', 'Capacidade_Unidade', 'Padrão', 'Excluir']
    df_cap = df_cap[['Item', 'Capacidade_Caixas', 'Capacidade_Unidade', 'Padrão']]
    df_cap = df_cap.apply(func=lambda val: pd.to_numeric(val, errors='coerce'))
    df_cap['Item'] = df_cap['Item'].astype(dtype='str')
    df_cap['Item'] = df_cap['Item'].fillna(0)

    df_saldo = pd.read_excel(nome_saldos)
    index = ['Item', 'Descrição Item',
             'Depósito', 'Q. Saldo Futuro']
    df_saldo = df_saldo[index]

    name_dep = ['DEPOSITO 2 - SHOW ROOM  MATRIZ', 'DEPOSITO 7 - VELOZ LOGISTICA', 'DEPOSITO 100 - CAIXAS PARDAS VELOZ',
                'DEPOSITO 115 - BLOQUEIOS/SEGREGADO PARA PEDIDOS', 'DEPOSITO 16 -RESERVA MADERO VELOZ',
                'DEPOSITO 20 - AVARIAS - VELOZ LOGISTICA', 'DEPOSITO 99 - BLOQUEIOS - IMPORTACAO']

    df_2 = df_saldo.query("Depósito == 'DEPOSITO 2 - SHOW ROOM  MATRIZ'")
    index2 = ['Item', 'Q. Saldo Futuro']
    df_2 = df_2[index2]
    df_2.columns = ['Item', 'Show_Room_2']
    df_2['Item'] = df_2['Item'].astype(dtype='str')
    df_2 = df_2.fillna(0)
    df_2['Show_Room_2'] = df_2['Show_Room_2'].apply(func=lambda val: pd.to_numeric(val))

    df_7 = df_saldo.query("Depósito == 'DEPOSITO 7 - VELOZ LOGISTICA'")
    index7 = ['Item', 'Descrição Item', 'Q. Saldo Futuro']
    df_7 = df_7[index7]
    df_7.columns = ['Item', 'Descrição_Item', 'Veloz_7']
    df_7['Item'] = df_7['Item'].astype(dtype='str')
    df_7 = df_7.fillna(0)

    df_100 = df_saldo.query("Depósito == 'DEPOSITO 100 - CAIXAS PARDAS VELOZ'")
    index100 = ['Item', 'Q. Saldo Futuro']
    df_100 = df_100[index100]
    df_100.columns = ['Item', 'Caixas_Pardas_Veloz_100']
    df_100['Item'] = df_100['Item'].astype(dtype='str')
    df_100 = df_100.fillna(0)

    df_115 = df_saldo.query("Depósito == 'DEPOSITO 115 - BLOQUEIOS/SEGREGADO PARA PEDIDOS'")
    index115 = ['Item', 'Q. Saldo Futuro']
    df_115 = df_115[index115]
    df_115.columns = ['Item', 'Bloqueios_115']
    df_115['Item'] = df_115['Item'].astype(dtype='str')
    df_115 = df_115.fillna(0)

    df_16 = df_saldo.query("Depósito == 'DEPOSITO 16 -RESERVA MADERO VELOZ'")
    index16 = ['Item', 'Q. Saldo Futuro']
    df_16 = df_16[index16]
    df_16.columns = ['Item', 'Reserva_Madero_16']
    df_16['Item'] = df_16['Item'].astype(dtype='str')
    df_16 = df_16.fillna(0)

    df_20 = df_saldo.query("Depósito == 'DEPOSITO 20 - AVARIAS - VELOZ LOGISTICA'")
    index20 = ['Item', 'Q. Saldo Futuro']
    df_20 = df_20[index20]
    df_20.columns = ['Item', 'Avarias_20']
    df_20['Item'] = df_20['Item'].astype(dtype='str')
    df_20 = df_20.fillna(0)

    df_99 = df_saldo.query("Depósito == 'DEPOSITO 99 - BLOQUEIOS - IMPORTACAO'")
    index99 = ['Item', 'Q. Saldo Futuro']
    df_99 = df_99[index99]
    df_99.columns = ['Item', 'Bloqueios_Importação_99']
    df_99['Item'] = df_99['Item'].astype(dtype='str')
    df_99 = df_99.fillna(0)

    df_guia = pd.merge(left=df_7, right=df_2, how='left', on='Item')
    df_guia = pd.merge(left=df_guia, right=df_100, how='left', on='Item')
    df_guia = pd.merge(left=df_guia, right=df_16, how='left', on='Item')
    df_guia = pd.merge(left=df_guia, right=df_115, how='left', on='Item')
    df_guia = pd.merge(left=df_guia, right=df_99, how='left', on='Item')
    df_guia = pd.merge(left=df_guia, right=df_20, how='left', on='Item')
    df_guia = pd.merge(left=df_guia, right=df_cap, how='left', on='Item')
    df_guia['Item'] = df_guia['Item'].apply(func=lambda val: pd.to_numeric(val))

    df_guia['Sugestão_Unidades'] = df_guia['Capacidade_Unidade'] - df_guia['Show_Room_2']
    df_guia['Sugestão_Caixas'] = df_guia['Sugestão_Unidades'] / df_guia['Padrão']
    df_guia['Sugestão_Caixas'] = df_guia['Sugestão_Caixas'].fillna(0)
    # df_guia['Sugestão_Caixas'] = df_guia['Sugestão_Caixas'].astype(dtype='str')
    df_guia['Sugestão_Caixas'] = df_guia['Sugestão_Caixas'].apply(func=lambda val: format_float(val))
    df_guia['Guia_Unidades'] = df_guia['Sugestão_Caixas'].apply(func=lambda val: '')
    df_guia['Guia_Caixas'] = df_guia['Sugestão_Caixas'].apply(func=lambda val: '')
    index_guia = ['Item', 'Descrição_Item', 'Veloz_7', 'Show_Room_2', 'Capacidade_Unidade', 'Sugestão_Unidades', 'Padrão',
                  'Sugestão_Caixas', 'Guia_Unidades', 'Guia_Caixas', 'Caixas_Pardas_Veloz_100', 'Bloqueios_115',
                  'Reserva_Madero_16', 'Avarias_20', 'Bloqueios_Importação_99']
    n_index_guia = ['Item', 'Descrição', 'Veloz-7', 'Show Room-2', 'Capac Un', 'Sugestão Un', 'Padrão',
                    'Sugestão Cx', 'Guia Cx', 'Guia Un', 'Cx Pardas-100', 'Bloq-115',
                    'Res Madero-16', 'Avarias-20', 'Bloq imp-99']

    df_guia = df_guia[index_guia]
    df_guia.columns = n_index_guia

    writer = pd.ExcelWriter(path=guia_file)
    df_guia.to_excel(excel_writer=writer, sheet_name='Sugestão_Guia', index=False)

    ws = writer.sheets['Sugestão_Guia']

    cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']
    dim = 13
    for z in cols:
        if z == 'A':
            ws.column_dimensions[z].width = 9

        if z == 'B':
            ws.column_dimensions[z].width = 40

        if z == 'C':
            ws.column_dimensions[z].width = dim

        if z == 'D':
            ws.column_dimensions[z].width = dim

        if z == 'E':
            ws.column_dimensions[z].width = dim

        if z == 'F':
            ws.column_dimensions[z].width = dim

        if z == 'G':
            ws.column_dimensions[z].width = dim

        if z == 'H':
            ws.column_dimensions[z].width = dim

        if z == 'I':
            ws.column_dimensions[z].width = dim

        if z == 'J':
            ws.column_dimensions[z].width = dim

        if z == 'K':
            ws.column_dimensions[z].width = dim

        if z == 'L':
            ws.column_dimensions[z].width = dim

        if z == 'M':
            ws.column_dimensions[z].width = dim

        if z == 'N':
            ws.column_dimensions[z].width = dim

        if z == 'O':
            ws.column_dimensions[z].width = dim

        if z == 'P':
            ws.column_dimensions[z].width = dim

    writer.save()

    import win32com.client as win32

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    # mail.To = 'patrick.paula@portoaporto.com.br'
    mail.To = 'rafael.felczak@portoaporto.com.br; faturamento2@portoaporto.com.br'
    mail.Subject = 'Robô guia'
    mail.Body = regra_cump()
    mail.Attachments.Add(guia_file)
    mail.Send()
