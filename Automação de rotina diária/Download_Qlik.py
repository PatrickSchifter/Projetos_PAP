from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui
import os
from datetime import datetime
import shutil
import pygetwindow
from Variaveis import meses

ac_day = datetime.today().strftime("%d")
ac_mon = datetime.today().strftime("%m")
ac_yea = datetime.today().strftime("%Y")

ac_day = int(ac_day)
ac_day = str(ac_day)
teste = 123
driver = webdriver.Chrome(
    executable_path=r"K:\CWB\Logistica\Rastreamento\Patrick\Automação\chromedriver_win32\chromedriver.exe")

# Abrir o Qlik
driver.get("https://portoaporto.us.qlikcloud.com/")


def wait_by_name(element):
    WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.NAME, element)))


def wait_by_id(element):
    WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.ID, element)))


wait_by_name("email")

time.sleep(1)

username = driver.find_element(By.NAME, "email")
password = driver.find_element(By.NAME, "password")
username.send_keys("rafael.felczak@portoaporto.com.br")
time.sleep(1)
password.send_keys("Rporto@3044")
next_step = driver.find_element(By.NAME, "submit")
next_step.click()

wait_by_id("item-open-button-612fade8a09189e80a049a2a")

time.sleep(2)

driver.get(
    "https://portoaporto.us.qlikcloud.com/sense/app/dc668b73-9cbe-49df-8031-5b2e1b23f699/sheet/0789deae-" +
    "e3f4-4be4-ad81-4b55b8e3d78c/state/analysis")

driver.get(
    "https://portoaporto.us.qlikcloud.com/sense/app/dc668b73-9cbe-49df-8031-5b2e1b23f699/sheet/0789deae-" +
    "e3f4-4be4-ad81-4b55b8e3d78c/state/analysis")

try:
    wait_by_id("60b01f31-6e88-46b8-82ae-3d8983f2267e-header-3")
except:
    pyautogui.press('f5')

time.sleep(1)
r_click = driver.find_element(By.ID, "60b01f31-6e88-46b8-82ae-3d8983f2267e-header-3")
time.sleep(1)
webdriver.ActionChains(driver).context_click(r_click).perform()

wait_by_id("child-menu")

time.sleep(1)
next_step = driver.find_element(By.ID, "child-menu")
next_step.click()

wait_by_id("download")

time.sleep(1)
next_step = driver.find_element(By.ID, "download")
next_step.click()

wait_by_id("download-data-tab")

time.sleep(1)
next_step = driver.find_element(By.ID, "download-data-tab")
next_step.click()

time.sleep(10)
pyautogui.click(795, 647)
time.sleep(5)

file_indir = True
while file_indir:
    path = list(os.listdir('C:/Users/patrick.paula/Downloads'))
    for file in path:
        if '.xlsx' in file:
            file_indir = False
            break
    for x in list(pygetwindow.getAllTitles()):
        if 'Qlik Sense' in x:
            title = x
    time.sleep(10)

    window = pygetwindow.getWindowsWithTitle(title)[0]
    window.activate()
    window.maximize()
    window.resizeTo(1366, 768)
    time.sleep(3)
    pyautogui.click(795, 647)
    time.sleep(3)
    pyautogui.click(946, 669)

time.sleep(1)

driver.close()

partial_name_1 = "Qlik Sense - Planilha de Alimentação de Datas - "
partial_name_2 = ".xlsx"
de = " de "
full_name_qlik = partial_name_1 + ac_day + de + meses[ac_mon] + de + ac_yea + partial_name_2
today = datetime.today().strftime("%d-%m-%Y")
p_source_path = "C:/Users/patrick.paula/Downloads/"
dest_path = "K:/CWB/Logistica/Rastreamento/Patrick/Storage/" + today

try:
    source_path_qlik = p_source_path + full_name_qlik
    dest_path_qlik = dest_path + "/" + full_name_qlik

    shutil.move(source_path_qlik, dest_path_qlik)
except:
    pass
