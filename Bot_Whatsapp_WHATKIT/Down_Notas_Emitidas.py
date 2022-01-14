import datetime
import pygetwindow
import pyautogui
import os
import time


def down_notas():
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

    pyautogui.write(datetime.datetime.today().strftime('%d/%m/%Y'))

    time.sleep(1)

    pyautogui.doubleClick(502, 239)

    pyautogui.write(datetime.datetime.today().strftime('%d/%m/%Y'))

    time.sleep(1)

    pyautogui.click(254, 138)

    time.sleep(2)

    pyautogui.click(314, 111)

    time.sleep(60)

    pyautogui.rightClick(470, 336)
    pyautogui.click(557, 383)
    time.sleep(2)
    pyautogui.click(862, 635)

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

    path = os.path.realpath('K:/CWB/Logistica/Rastreamento/Automacoes/WhatsExpedição/Notas')
    os.startfile(path)
    time.sleep(2)
    pyautogui.click(552, 307)

    with pyautogui.hold('ctrlright'):
        pyautogui.press('v')

    time.sleep(5)

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
