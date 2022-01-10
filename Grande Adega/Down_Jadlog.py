from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC



def wait_by_name(element):
    element = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.NAME, element)))


def wait_by_id(element):
    element = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, element)))

driver = webdriver.Chrome(
    executable_path=r"K:\CWB\Logistica\Rastreamento\Patrick\Automação\chromedriver_win32\chromedriver.exe")
driver.get('https://www.jadlog.com.br/portalcliente/pages/login.jad')

wait_by_id('id_usuario')
user_id = driver.find_element_by_id('id_usuario')
user_ps = driver.find_element_by_id('id_senha')
user_id.send_keys('jadlog.grandeadega')
user_ps.send_keys('GrandeAdega21')
submit = driver.find_element_by_name('j_idt15')
submit.click()
wait_by_id('j_idt4')
driver.get('https://www.jadlog.com.br/portalcliente/pages/relatorio_entrega/index.jad')