
### Automatizaci[o]n con Selenium TSICAM
### By Marlhen Estrada Dic 2021

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook

import time

#### Ingresamos url, usuario y la contrasena.

username = "mestradav@southernperu.com.pe"
password = "Nuevaera202112"
url = "https://tsicam.southernperu.com.pe/intranet/documentos/area/revision"

#### Arrancamos lo automatizado

driver=webdriver.Chrome(executable_path=r"C:\driver\chromedriver.exe")
driver.get(url)
driver.maximize_window()
driver.find_element_by_name("username").send_keys(username)
driver.find_element_by_name("password").send_keys(password)
driver.find_element_by_css_selector("input.btn.btn-primary").click()
time.sleep(3)

#### Cargamos el excel donde estan los DNI[s]

filesheet = "D:\\Mico\\Cursos\\Python\\Selenium\\prueba.xlsx"
wb = load_workbook(filesheet)
hojas = wb.get_sheet_names()
print(hojas)
nombres = wb.get_sheet_by_name('Sheet1')
wb.close()

#### Hacemos For para la cantidad de ingresos

for i in range(1,6):
    dni, nomb = nombres[f'A{i}:B{i}'][0]
    print(dni.value, nomb.value)
    time.sleep(3)
    driver.find_element_by_name("nombre").send_keys(dni.value)
    driver.find_element_by_name("nombre").send_keys(Keys.ENTER)
    time.sleep(3)


    driver.find_element_by_xpath("//a[@title='Reagendar Induccion y Orientacion Básica de Seguridad - Toquepala']").click()
    driver.find_element_by_name("fechaSolicitud").clear()
    time.sleep(3)
    
    #### La fecha es importante

    driver.find_element_by_name("fechaSolicitud").send_keys("11/05/2022")
    time.sleep(3)
    driver.find_element_by_xpath("(//a[@id='save-modal'])[1]").click()
    time.sleep(3)
    driver.find_element_by_css_selector(".fa.fa-check").click()
    time.sleep(3)
    driver.find_element_by_css_selector("#id_anexo").send_keys("D:\\Mico\\Cursos\\Python\\Selenium\\"+ str(dni.value) +".pdf")
    time.sleep(3)
    
    #driver.find_element_by_css_selector("#save-modal").click()
    #time.sleep(3)
    driver.find_element_by_xpath("(//a[@class='modal-close waves-effect waves-green btn-flat'][normalize-space()='Cancelar'])[1]").click()
    time.sleep(3)
    driver.find_element_by_name("nombre").clear()
    time.sleep(3)


