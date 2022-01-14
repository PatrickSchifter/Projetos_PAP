import webbrowser as web
import pyautogui
import time
import os
from config import path_dir_tod, today
import pandas as pd

web.open(f"https://web.whatsapp.com/send?phone=+5541991912238")

time.sleep(20)
pyautogui.moveTo(665, 645)
pyautogui.dragTo(706, 50, 3, button='left')

pyautogui.keyDown('ctrl')
pyautogui.press('c')
pyautogui.keyUp('ctrl')
file_name = path_dir_tod + f'/Entregas {today}.txt'

try:
    os.startfile(file_name)
except FileNotFoundError:
    file = open(file_name, 'w')
    os.startfile(file_name)

time.sleep(3)

pyautogui.keyDown('ctrl')
pyautogui.press('v')
pyautogui.keyUp('ctrl')
time.sleep(1)
pyautogui.keyDown('ctrl')
pyautogui.press('s')
pyautogui.keyUp('ctrl')
time.sleep(1)
pyautogui.press('enter')
pyautogui.press('left')
pyautogui.press('enter')
pyautogui.keyDown('alt')
pyautogui.press('f4')
pyautogui.keyUp('alt')
time.sleep(2)
pyautogui.press('enter')

list_num = []
list_dat = []
with open(file_name, 'r') as file:
    for row in file:
        row = row.replace('\n', '')
        row = row.replace(' ', '')
        if len(row) == 6 and row[0].isnumeric():
            list_num.append(row)
            list_dat.append(today)


df_list = {'Número': list_num, 'D_Entrega': list_dat}
df_ent = pd.DataFrame(df_list)
df_ent = df_ent.drop_duplicates(subset=['Número'], keep='last')

print(df_ent)

