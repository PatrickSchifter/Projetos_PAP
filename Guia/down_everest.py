import pyautogui
import time
import os
from datetime import datetime
import pygetwindow

path_f = r'K:/CWB/Logistica/Rastreamento/Automacoes/Guia/Arquivo'
nome_saldos = path_f + '/Saldo_estoque ' + datetime.today().strftime('%d-%m-%Y') + '.xlsx'

def down_notas_e():
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
    pyautogui.click(773, 315)
    time.sleep(15)
    pyautogui.click(93, 631)
    time.sleep(1)
    pyautogui.click(32, 138)
    time.sleep(1)
    pyautogui.click(108, 155)
    time.sleep(5)
    pyautogui.click(575, 207)
    time.sleep(1)
    pyautogui.doubleClick(517, 293)
    time.sleep(1)
    pyautogui.click(294, 285)
    pyautogui.write(datetime.today().strftime('%d/%m/%Y'))
    time.sleep(1)
    pyautogui.click(298, 437)
    time.sleep(1)
    pyautogui.click(275, 468)
    time.sleep(1)
    pyautogui.click(255, 147)
    time.sleep(60)

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



    path = os.path.realpath(path_f)
    os.startfile(path_f)
    time.sleep(2)
    pyautogui.click(552, 307)

    with pyautogui.hold('ctrlright'):
        pyautogui.press('v')

    time.sleep(5)

    itens_indir = list(os.listdir(path_f))

    for item in itens_indir:
        if "Saldos" in item:
            dir = path_f + '/' + item
            os.rename(dir, nome_saldos)

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
