import pygetwindow
import pyautogui
import os
import time
from config import path_dir_tod, dataf, path_r, ano_atual, partial_path, p_source_path

# Criação de diretório na rede
try:
    os.mkdir(path_dir_tod)
    print('Criado diretório na rede')
except FileNotFoundError:
        os.mkdir(path_r + ano_atual)
        os.mkdir(partial_path)
        os.mkdir(path_dir_tod)
        print('Criado diretório na rede')
except FileExistsError:
    pass

try:
    for itens in list(os.listdir(p_source_path)):
        os.remove(p_source_path + itens)
except:
    print('Limpeza da pasta downloads')
# print('Deu algum erro ao criar o diretório')

# Abertura do Everest

a_remota = "C:/Users/patrick.paula/Desktop/EVEREST_3.0.rdp"
a_remota = os.path.realpath(a_remota)
os.startfile(a_remota)
time.sleep(5)

pyautogui.write('TW0a5yR4')

pyautogui.press('Enter')
time.sleep(1)
pyautogui.write('TW0a5yR4')

pyautogui.press('Enter')

time.sleep(30)
pyautogui.write('123456')
time.sleep(2)
pyautogui.press('Enter')

time.sleep(20)

pyautogui.click(905, 326)

time.sleep(20)

pyautogui.click(49, 600)

time.sleep(1)

pyautogui.click(69, 114)

time.sleep(5)

pyautogui.click(291, 238)

time.sleep(1)

pyautogui.click(358, 239)

time.sleep(1)

pyautogui.doubleClick(398, 239)

pyautogui.write(dataf())

time.sleep(1)

pyautogui.doubleClick(502, 239)

pyautogui.write(dataf())

time.sleep(1)

pyautogui.click(254, 138)

time.sleep(2)

pyautogui.click(314, 111)

time.sleep(2)

import Download_Qlik
import Download_Plan_Urano
import N_Coleta_Veloz

######################################################################################

# Localiza a janela do Everest na barra de tarefas e maximiza.

for x in list(pygetwindow.getAllTitles()):
    if "EVEREST" in x:
        title = x

try:
    window = pygetwindow.getWindowsWithTitle(title)[0]
    window.activate()
    window.maximize()
    window.resizeTo(1366, 768)
except:
    pass

import Baixar_Notas_Emitidas
import Concat_plan

for x in list(pygetwindow.getAllTitles()):
    if "EVEREST" in x:
        title = x

window = pygetwindow.getWindowsWithTitle(title)[0]
window.activate()
window.maximize()
window.resizeTo(1366, 768)

pyautogui.click(1234, 106)
time.sleep(1)
pyautogui.click(1240, 48)
time.sleep(1)
pyautogui.click(1256, 177)
time.sleep(1)
pyautogui.click(1240, 48)
time.sleep(1)
pyautogui.click(1267, 199)
time.sleep(1)

import DF_Qlik
