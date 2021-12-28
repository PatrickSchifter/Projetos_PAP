from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime
from datetime import date
import os
import shutil
import pyautogui
from Variaveis import source_path_urano, dest_path_urano


def wait_by_name(element):
    element = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.NAME, element)))


def wait_by_id(element):
    element = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, element)))

    # driver Chrome


driver = webdriver.Chrome(
    executable_path=r"K:\CWB\Logistica\Rastreamento\Patrick\Automação\chromedriver_win32\chromedriver.exe")

driver.get("https://platform.senior.com.br/")

wait_by_id("username-input-field")
username_ura = driver.find_element_by_id("username-input-field")
username_ura.send_keys("porto@uranolog.com.br")
password_ura = driver.find_element_by_id("password-input-field")
password_ura.send_keys("Senha1!")

next_step = driver.find_element_by_id("loginbtn")
next_step.click()

driver.get(
    'https://platform.senior.com.br/senior-x/#/Gest%C3%A3o%20Log%C3%ADstica/2/69a7e7e3-aaaa-4eb5-ac60-c2514b0c010e%2FGest%C3%A3o%20de%20Transporte%2FPortal%20do%20cliente%2FRelat%C3%B3rio%20de%20situa%C3%A7%C3%A3o%20de%20transporte?category=frame&link=https:%2F%2Fplatform.senior.com.br%2Flogistica-documentos%2Ftms%2Fdocumentos-frontend%2F%23%2Fdocumentos&r=0')
time.sleep(1)
driver.get(
    'https://platform.senior.com.br/senior-x/#/Gest%C3%A3o%20Log%C3%ADstica/2/69a7e7e3-aaaa-4eb5-ac60-c2514b0c010e%2FGest%C3%A3o%20de%20Transporte%2FPortal%20do%20cliente%2FRelat%C3%B3rio%20de%20situa%C3%A7%C3%A3o%20de%20transporte?category=frame&link=https:%2F%2Fplatform.senior.com.br%2Flogistica-documentos%2Ftms%2Fdocumentos-frontend%2F%23%2Fdocumentos&r=0')

hoje = datetime.today()

time.sleep(3)
pyautogui.click(138, 667)
time.sleep(1)
pyautogui.write(date.fromordinal(hoje.toordinal() - 30).strftime('%d-%m-%Y'))
time.sleep(1)
pyautogui.click(442, 666)
pyautogui.write(datetime.today().strftime('%d-%m-%Y'))
pyautogui.keyDown("PageDown")
time.sleep(2)
pyautogui.click(182, 499)
time.sleep(8)

pyautogui.keyDown("PageUp")
time.sleep(1)
pyautogui.click(188, 490)
time.sleep(1)
pyautogui.click(728, 314)
time.sleep(10)

pyautogui.click(46, 285)
time.sleep(1)
pyautogui.click(186, 288)

time.sleep(10)
driver.close()

path = list(os.listdir("C:/Users/patrick.paula/Downloads"))

full_name_urano = "C:/Users/patrick.paula/Downloads/" + "Urano " + str(
    datetime.today().strftime('%d-%m-%Y')) + ".CSV"

for file in path:
    if ".CSV" in file:
        file_name = "C:/Users/patrick.paula/Downloads/" + file
        os.rename(file_name, full_name_urano)

full_name_urano = "Urano " + str(datetime.today().strftime('%d-%m-%Y')) + ".CSV"

today = datetime.today().strftime("%d-%m-%Y")
p_source_path = "C:/Users/patrick.paula/Downloads/"
dest_path = "K:/CWB/Logistica/Rastreamento/Patrick/Storage/" + today

try:

    source_path_urano = p_source_path + full_name_urano
    dest_path_urano = dest_path + "/" + full_name_urano

    shutil.move(source_path_urano, dest_path_urano)
except:
    driver = webdriver.Chrome(
        executable_path=r"K:\CWB\Logistica\Rastreamento\Patrick\Automação\chromedriver_win32\chromedriver.exe")

    driver.get("https://platform.senior.com.br/")

    wait_by_id("username-input-field")
    username_ura = driver.find_element_by_id("username-input-field")
    username_ura.send_keys("porto@uranolog.com.br")
    password_ura = driver.find_element_by_id("password-input-field")
    password_ura.send_keys("Senha1!")

    next_step = driver.find_element_by_id("loginbtn")
    next_step.click()

    driver.get(
        'https://platform.senior.com.br/senior-x/#/Gest%C3%A3o%20Log%C3%ADstica/2/69a7e7e3-aaaa-4eb5-ac60-c2514b0c010e%2FGest%C3%A3o%20de%20Transporte%2FPortal%20do%20cliente%2FRelat%C3%B3rio%20de%20situa%C3%A7%C3%A3o%20de%20transporte?category=frame&link=https:%2F%2Fplatform.senior.com.br%2Flogistica-documentos%2Ftms%2Fdocumentos-frontend%2F%23%2Fdocumentos&r=0')
    time.sleep(1)
    driver.get(
        'https://platform.senior.com.br/senior-x/#/Gest%C3%A3o%20Log%C3%ADstica/2/69a7e7e3-aaaa-4eb5-ac60-c2514b0c010e%2FGest%C3%A3o%20de%20Transporte%2FPortal%20do%20cliente%2FRelat%C3%B3rio%20de%20situa%C3%A7%C3%A3o%20de%20transporte?category=frame&link=https:%2F%2Fplatform.senior.com.br%2Flogistica-documentos%2Ftms%2Fdocumentos-frontend%2F%23%2Fdocumentos&r=0')

    hoje = datetime.today()

    time.sleep(3)
    pyautogui.click(138, 667)
    time.sleep(1)
    pyautogui.write(date.fromordinal(hoje.toordinal() - 30).strftime('%d-%m-%Y'))
    time.sleep(1)
    pyautogui.click(442, 666)
    pyautogui.write(datetime.today().strftime('%d-%m-%Y'))
    pyautogui.keyDown("PageDown")
    time.sleep(2)
    pyautogui.click(182, 499)
    time.sleep(8)

    pyautogui.keyDown("PageUp")
    time.sleep(1)
    pyautogui.click(188, 490)
    time.sleep(1)
    pyautogui.click(728, 314)
    time.sleep(10)

    pyautogui.click(46, 285)
    time.sleep(1)
    pyautogui.click(186, 288)

    time.sleep(10)
    driver.close()

    path = list(os.listdir("C:/Users/patrick.paula/Downloads"))

    full_name_urano = "C:/Users/patrick.paula/Downloads/" + "Urano " + str(
        datetime.today().strftime('%d-%m-%Y')) + ".CSV"

    for file in path:
        if ".CSV" in file:
            file_name = "C:/Users/patrick.paula/Downloads/" + file
            os.rename(file_name, full_name_urano)

    shutil.move(source_path_urano, dest_path_urano)
