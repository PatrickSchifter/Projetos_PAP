import Variaveis
import pyautogui
import os
import time
from Variaveis import dataf
time.sleep(5)
pyautogui.rightClick(470, 336)
pyautogui.click(557, 383)
time.sleep(2)
pyautogui.click(862, 440)

time.sleep(5)

pyautogui.click(268, 216)
time.sleep(1)
pyautogui.rightClick(470, 336)
pyautogui.rightClick(571, 350)
pyautogui.rightClick(843, 374)
time.sleep(5)
pyautogui.press('enter')
time.sleep(5)

pyautogui.rightClick(467, 191)
pyautogui.click(516, 226)

time.sleep(5)

path = os.path.realpath(Variaveis.path_dir_tod)
os.startfile(Variaveis.path_dir_tod)
time.sleep(2)
pyautogui.click(552, 307)

with pyautogui.hold('ctrlright'):
    pyautogui.press('v')

time.sleep(5)

itens_indir = list(os.listdir(Variaveis.path_dir_tod))
name_notas = Variaveis.path_dir_tod + '/Notas_emitidas ' + dataf() + '.xlsx'

for item in itens_indir:
    if "Análise" in item:
        dir = Variaveis.path_dir_tod + '/' + item
        os.rename(dir, name_notas)
