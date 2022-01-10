import shutil
import time
from selenium import webdriver
import pyautogui
import os
from datetime import datetime
from Variaveis import dest_path, ano_atual

path_driver = "K:\CWB\Logistica\Rastreamento\Patrick\Automação\chromedriver_win32\chromedriver.exe"

driver = webdriver.Chrome(executable_path=path_driver)

link_gdrive = 'https://drive.google.com/drive/folders/1_wfu6sA9Jplfa2uaytx8dXPqIXYMQ9o8?usp=sharing'

driver.get(link_gdrive)
path_file = f'C:/Users/patrick.paula/Downloads/CONTROLE DE PEDIDOS PAP {ano_atual}.xlsx'

time.sleep(3)
pyautogui.click(332, 332)
time.sleep(10)

file_indir = True
while file_indir:
    path = list(os.listdir('C:/Users/patrick.paula/Downloads'))
    for file in path:
        if 'CONTROLE DE PEDIDOS' in file:
            time.sleep(1)
            file_indir = False
            break

    time.sleep(5)
    pyautogui.click(332, 332)

shutil.move(path_file, dest_path)

driver.close()
