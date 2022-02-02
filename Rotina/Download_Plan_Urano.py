def download_plan_urano():
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
    from Projetos.Rotina.config import source_path_urano, dest_path_urano, full_name_urano, dest_path, open_window

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
    time.sleep(3)
    open_window('Senior')

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
    pyautogui.click(145, 650)
    time.sleep(1)
    pyautogui.write(date.fromordinal(hoje.toordinal() - 30).strftime('%d-%m-%Y'))
    time.sleep(1)
    pyautogui.click(554, 649)
    pyautogui.write(datetime.today().strftime('%d-%m-%Y'))
    pyautogui.keyDown("PageDown")
    time.sleep(2)
    pyautogui.click(167, 453)
    time.sleep(8)

    pyautogui.keyDown("PageUp")
    time.sleep(1)
    pyautogui.click(197, 464)
    time.sleep(1)
    pyautogui.click(1023, 299)
    time.sleep(10)

    pyautogui.click(30, 267)
    time.sleep(1)
    pyautogui.click(166, 271)

    time.sleep(10)
    driver.close()

    path = list(os.listdir("C:/Users/patrick.paula/Downloads"))

    for file in path:
        if ".CSV" in file:
            file_name = "C:/Users/patrick.paula/Downloads/" + file
            shutil.move(file_name, dest_path + '/' + file)

    for file in list(os.listdir(dest_path)):
        if ".CSV" in file:
            file_name = dest_path + '/' + file
            os.rename(file_name, dest_path_urano)

