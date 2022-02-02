from Projetos.Rotina import config
import pyautogui
import os
import time
from Projetos.Rotina.config import dataf
import pygetwindow

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

path = os.path.realpath(config.path_dir_tod)
os.startfile(config.path_dir_tod)
time.sleep(2)
pyautogui.click(552, 307)

with pyautogui.hold('ctrlright'):
    pyautogui.press('v')

time.sleep(5)

itens_indir = list(os.listdir(config.path_dir_tod))
name_notas = config.path_dir_tod + '/Notas_emitidas ' + dataf() + '.xlsx'

for item in itens_indir:
    if "Análise" in item:
        dir = config.path_dir_tod + '/' + item
        os.rename(dir, name_notas)

for x in list(pygetwindow.getAllTitles()):
    if "EVEREST" in x:
        title = x

window = pygetwindow.getWindowsWithTitle(title)[0]
window.activate()
window.maximize()
window.resizeTo(1366, 768)