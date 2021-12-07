import pandas as pd
from Variaveis import partial_path, all_index, dest_file, dia_atual, ano_atual, mes_atual, today, dataf, dia_semana
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.styles import NamedStyle, Font, Border, Side, PatternFill
from openpyxl import Workbook


file = r'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/MONITORAMENTO 2021 v.1.2.xlsx'  # Planilha de Monitoramento
n_file = partial_path + today + '/Notas_emitidas ' + dataf() + '.xlsx'  # Notas do dia anterior

itens_comp = []
sheet_names = []

path_plans = 'K:/CWB/Logistica/Rastreamento/Controle_Monitoramento/Dados importantes'

sheet_name = 'Clientes com Agendamento'
df = pd.read_excel(file, sheet_name=sheet_name, engine='openpyxl')
sheet_names.append(sheet_name)
itens_comp.append(df)

sheet_name = 'Contato Transportadoras'
df = pd.read_excel(file, sheet_name=sheet_name, engine='openpyxl')
sheet_names.append(sheet_name)
itens_comp.append(df)

df_n = pd.read_excel(n_file, engine='openpyxl')

meses = {"1": "JANEIRO",
         "2": "FEVEREIRO",
         "3": "MARÇO",
         "4": "ABRIL",
         "5": "MAIO",
         "6": "JUNHO",
         "7": "JULHO",
         "8": "AGOSTO",
         "9": "SETEMBRO",
         "10": "OUTUBRO",
         "11": "NOVEMBRO",
         "12": "DEZEMBRO"}
dfs = []

try:
    for dfss in range(1, 13):
        dfss = str(dfss)
        try:
            df = pd.read_excel(file, engine='openpyxl', sheet_name=meses[dfss] + '-' + ano_atual)
            df.columns = all_index
        except ValueError:
            pass

        dfs.append(df)
except IndexError:
    pass

index = ['Data-De-Coleta', 'Previsão-Entrega', 'D_Entrega', 'Agendamento', 'Lead-Time',
         'Dias-Para-Entrega', 'Resumo', 'Enviar-Email']

df_n.columns = ['Número', 'N.Pré-Nota', 'Emissão', 'Fantasia-Destinatário', 'Cidade-Destinatário', 'Uf',
                'Natureza-Fiscal', 'Situação-Fiscal', 'Descrição-Do-Depósito', 'Fantasia_Do_Transportador',
                'Fantasia-Comissionado']

for x in index:
    df_n[x] = ''

if int(dia_atual) == 1:
    if int(dia_atual) == 2 and dia_semana == 0:
        dfs[int(mes_atual) - 2] = pd.concat([dfs[int(mes_atual) - 2], df_n],
                                            keys=all_index)

    if int(dia_atual) == 3 and dia_semana == 0:
        dfs[int(mes_atual) - 2] = pd.concat([dfs[int(mes_atual) - 2], df_n],
                                            keys=all_index)

    dfs[int(mes_atual) - 2] = pd.concat([dfs[int(mes_atual) - 2], df_n],
                                        keys=all_index)
else:
    dfs[int(mes_atual) - 1] = pd.concat([dfs[int(mes_atual) - 1], df_n],
                                        keys=all_index)


from SSW import name_m_ant, name_m_atual, search_data

ssw_mes = search_data(name_m_atual)
ssw_mes_m = search_data(name_m_ant)

ssw_mes = ssw_mes[['Número', 'Ult_Status']]
ssw_mes_m = ssw_mes_m[['Número', 'Ult_Status']]

dfs[int(mes_atual) - 2] = pd.merge(left=dfs[int(mes_atual) - 2], right=ssw_mes_m, how='left')
dfs[int(mes_atual) - 1] = pd.merge(left=dfs[int(mes_atual) - 1], right=ssw_mes, how='left')

dfs[int(mes_atual) - 2] = dfs[int(mes_atual) - 2][['Número', 'N.Pré-Nota', 'Emissão', 'Fantasia-Destinatário', 'Cidade-Destinatário', 'Uf',
             'Natureza-Fiscal', 'Situação-Fiscal', 'Descrição-Do-Depósito', 'Fantasia_Do_Transportador',
             'Fantasia-Comissionado', 'Data-De-Coleta', 'Previsão-Entrega', 'D_Entrega', 'Agendamento',
             'Lead-Time',
             'Dias-Para-Entrega', 'Resumo', 'Ult_Status']]
dfs[int(mes_atual) - 2].columns = all_index

dfs[int(mes_atual) - 1] = dfs[int(mes_atual) - 1][['Número', 'N.Pré-Nota', 'Emissão', 'Fantasia-Destinatário', 'Cidade-Destinatário', 'Uf',
             'Natureza-Fiscal', 'Situação-Fiscal', 'Descrição-Do-Depósito', 'Fantasia_Do_Transportador',
             'Fantasia-Comissionado', 'Data-De-Coleta', 'Previsão-Entrega', 'D_Entrega', 'Agendamento',
             'Lead-Time',
             'Dias-Para-Entrega', 'Resumo', 'Ult_Status']]

dfs[int(mes_atual) - 1].columns = all_index

cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']
dates = ['C', 'L', 'M', 'N', 'O']

book = Workbook()

index_style = NamedStyle(name='index_style',
                         fill=PatternFill(start_color='00C0C0C0', end_color='00C0C0C0', fill_type='solid'))
index_style.font = Font(size=9, bold=True)
bd_index = Side(style='thin', color='000000')
index_style.border = Border(left=bd_index, right=bd_index, top=bd_index, bottom=bd_index)
index_style.alignment = Alignment(horizontal='center')
book.add_named_style(index_style)

date_format = NamedStyle(name='date_format', number_format='DD/MM/YYYY', font=Font(color='000000'))
date_format.font = Font(size=8)
date_format.border = Border(left=bd_index, right=bd_index, top=bd_index, bottom=bd_index)
date_format.alignment = Alignment(horizontal='center')
book.add_named_style(date_format)

default_style = NamedStyle(name='default_style', font=Font(color='000000'))
default_style.font = Font(size=8)
default_style.border = Border(left=bd_index, right=bd_index, top=bd_index, bottom=bd_index)
default_style.alignment = Alignment(horizontal='center')
book.add_named_style(default_style)

try:
    for x in range(1, len(dfs) + 1):
        sheet_name = meses[str(x)] + '-' + ano_atual
        print(sheet_name + ' incluso na planilha')

        if x == 1:
            date = NamedStyle(name='date', number_format='DD/MM/YYYY', font=Font(color='000000'))
            date.font = Font(size=8)
            date.border = Border(left=bd_index, right=bd_index, top=bd_index, bottom=bd_index)
            date.alignment = Alignment(horizontal='center')
            book.add_named_style(date)

        if x > 1:
            book = load_workbook(filename=dest_file)
            writer = pd.ExcelWriter(path=dest_file, engine='openpyxl', datetime_format='DD/MM/YYYY',
                                    date_format='dd/mm/yyyy')
            writer.book = book

        if x == 1:
            writer = pd.ExcelWriter(path=dest_file,
                                    engine='openpyxl', index=False)

        dfs[x - 1] = dfs[x - 1][all_index]

        dfs[x - 1].to_excel(excel_writer=writer, sheet_name=sheet_name, index=False)

        ws = writer.sheets[sheet_name]

        for z in cols:
            for cell in ws[z]:
                cell.style = default_style

            if z == 'A' or 'B':
                ws.column_dimensions[z].width = 9

            if z in dates:
                ws.column_dimensions[z].width = 13
                for cell in ws[z]:
                    cell.style = date_format

                if x >= 2:
                    for cell in ws[z]:
                        cell.style = date

            if z == 'D':
                ws.column_dimensions[z].width = 33

            if z == 'E':
                ws.column_dimensions[z].width = 30

            if z == 'F':
                ws.column_dimensions[z].width = 4

            if z == 'G':
                ws.column_dimensions[z].width = 24

            if z == 'H':
                ws.column_dimensions[z].width = 13

            if z == 'I':
                ws.column_dimensions[z].width = 37

            if z == 'J':
                ws.column_dimensions[z].width = 25

            if z == 'K':
                ws.column_dimensions[z].width = 25

            if z == 'Q':
                ws.column_dimensions[z].width = 14

            if z == 'S':
                ws.column_dimensions[z].width = 30

            # Index
            col = z + '1'
            ws[col].style = index_style

        writer.save()
except IndexError:
    writer.save()

# Clientes com Agendamento

book = load_workbook(filename=dest_file)
writer = pd.ExcelWriter(path=dest_file, engine='openpyxl', datetime_format='DD/MM/YYYY', date_format='dd/mm/yyyy')
writer.book = book
itens_comp[0].to_excel(excel_writer=writer, sheet_name=sheet_names[0], index=False)

ws = writer.sheets[sheet_names[0]]

style2 = NamedStyle(name='style2', font=Font(size=10, color='000000'))
style2.border = Border(left=bd_index, right=bd_index, top=bd_index, bottom=bd_index)
style2.alignment = Alignment(horizontal='center', vertical='center')
book.add_named_style(style2)

for c in cols[0:3]:
    for cell in ws[c]:
        cell.style = style2

    if c == 'A':
        ws.column_dimensions[c].width = 40

    if c == 'B':
        ws.column_dimensions[c].width = 70

    if c == 'C':
        ws.column_dimensions[c].width = 50

    # Index
    col = c + '1'
    ws[col].style = index_style

ws.merge_cells('A4:A5')
ws.merge_cells('A6:A11')
ws.merge_cells('A29:A30')

print(sheet_names[0] + ' Incluso na Planilha')

writer.save()

# Contato Transportadora

book = load_workbook(filename=dest_file)
writer = pd.ExcelWriter(path=dest_file, engine='openpyxl', datetime_format='DD/MM/YYYY', date_format='dd/mm/yyyy')
writer.book = book
itens_comp[1].to_excel(excel_writer=writer, sheet_name=sheet_names[1], index=False)

ws = writer.sheets[sheet_names[1]]

style3 = NamedStyle(name='style3', font=Font(size=10, color='000000'))
style3.border = Border(left=bd_index, right=bd_index, top=bd_index, bottom=bd_index)
style3.alignment = Alignment(horizontal='center', vertical='center')
book.add_named_style(style3)

for c in cols[0:5]:
    for cell in ws[c]:
        cell.style = style3

    if c == 'A':
        ws.column_dimensions[c].width = 18

    if c == 'B':
        ws.column_dimensions[c].width = 25

    if c == 'C':
        ws.column_dimensions[c].width = 40

    if c == 'D':
        ws.column_dimensions[c].width = 45

    if c == 'E':
        ws.column_dimensions[c].width = 65
    # Index
    col = c + '1'
    ws[col].style = index_style

ws.merge_cells('A2:A6')
ws.merge_cells('A7:A8')
ws.merge_cells('A9:A13')
ws.merge_cells('A14:A18')
ws.merge_cells('A19:A23')
ws.merge_cells('A24:A29')
ws.merge_cells('A30:A32')
ws.merge_cells('A33:A35')
ws.merge_cells('A36:A38')
ws.merge_cells('A39:A44')
ws.merge_cells('A45:A47')
ws.merge_cells('A48:A56')
ws.merge_cells('A57:A64')
ws.merge_cells('A65:A70')
ws.merge_cells('A72:A75')
ws.merge_cells('A76:A80')
ws.merge_cells('A81:A82')
ws.merge_cells('A83:A89')
ws.merge_cells('A90:A95')
ws.merge_cells('A96:A97')
ws.merge_cells('A98:A104')
ws.merge_cells('A105:A106')
ws.merge_cells('B2:B6')
ws.merge_cells('B7:B8')
ws.merge_cells('B9:B13')
ws.merge_cells('B14:B18')
ws.merge_cells('B19:B23')
ws.merge_cells('B24:B29')
ws.merge_cells('B30:B32')
ws.merge_cells('B33:B35')
ws.merge_cells('B36:B38')
ws.merge_cells('B39:B44')
ws.merge_cells('B45:B47')
ws.merge_cells('B48:B56')
ws.merge_cells('B57:B64')
ws.merge_cells('B65:B70')
ws.merge_cells('B72:B75')
ws.merge_cells('B76:B80')
ws.merge_cells('B81:B82')
ws.merge_cells('B83:B89')
ws.merge_cells('B90:B95')
ws.merge_cells('B96:B97')
ws.merge_cells('B98:B104')
ws.merge_cells('B105:B106')

print(sheet_names[1] + ' Incluso na Planilha')

writer.save()
