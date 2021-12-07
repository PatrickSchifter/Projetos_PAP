import pygetwindow
import pyautogui
import os
import time
from pythonProject1.Arquivo.Data_anterior import data
from Variaveis import path_dir_tod

# Criação de diretório na rede

try:
    os.mkdir(path_dir_tod)
except:
    pass

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

pyautogui.write(data())

time.sleep(1)

pyautogui.doubleClick(502, 239)

pyautogui.write(data())

time.sleep(1)

pyautogui.click(254, 138)

time.sleep(2)

pyautogui.click(314, 111)

time.sleep(2)

######################################################################################

# Localiza a janela do Everest na barra de tarefas e maximiza.

for x in list(pygetwindow.getAllTitles()):
    if "EVEREST" in x:
        title = x

window = pygetwindow.getWindowsWithTitle(title)[0]
window.activate()
window.maximize()
window.resizeTo(1366, 768)

# import SSW
# import Trat_dados_ssw
