### Automatizaci[o]n con Selenium TSICAM
### By Marlhen Estrada Dic 2022 Rev 0.2
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
time.sleep(1)

#### Cargamos el excel donde estan los DNI[s]

filesheet = "C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\excel.xlsx"
wb = load_workbook(filesheet)
hojas = wb.get_sheet_names()
print(hojas)
nombres = wb.get_sheet_by_name('Sheet1')
wb.close()

#### Hacemos For para la cantidad de ingresos
# dni, nomb = A y B son las columnas  y [0] empieza en la primera columna


for i in range(1,20):
    dni, nomb = nombres[f'A{i}:B{i}'][0]
    print(dni.value, nomb.value)
    time.sleep(2)
    driver.find_element_by_name("nombre").send_keys(dni.value)
    driver.find_element_by_name("nombre").send_keys(Keys.ENTER)
    time.sleep(2)

    try:
    ## Verifica si el selector existe.
        item = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,"//a[@title='Reagendar Induccion y Orientacion Básica de Seguridad - Cuajone']")))
    #do something with the item...
        driver.find_element_by_xpath("//a[@title='Reagendar Induccion y Orientacion Básica de Seguridad - Cuajone']").click()
        driver.find_element_by_name("fechaSolicitud").clear()
        time.sleep(2)
        driver.find_element_by_name("fechaSolicitud").send_keys("21/05/2022")
        time.sleep(2)
        driver.find_element_by_xpath("(//a[@id='save-modal'])[1]").click()
        time.sleep(2)
        driver.find_element_by_css_selector(".fa.fa-check").click()
        time.sleep(2)

        #colocar la nota
        driver.find_element_by_xpath("(//input[@id='id_calificacion'])[1]").send_keys("17")
        time.sleep(2)

        #Sube el anexo 4

        driver.find_element_by_css_selector("#id_anexo").send_keys("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\"+ str(dni.value) +".pdf")
        time.sleep(2)
        driver.find_element_by_css_selector("#save-modal").click()
        time.sleep(2)
        
        #driver.find_element_by_xpath("(//a[@class='modal-close waves-effect waves-green btn-flat'][normalize-space()='Cancelar'])[1]").click()
        #time.sleep(1)
        
        driver.find_element_by_name("nombre").clear()
        time.sleep(2)

    except TimeoutException as e:
        print("Ya esta activado")
        pass
        time.sleep(2)

    driver.find_element_by_name("nombre").clear()
    time.sleep(2)

    #Repositorio para cargar los archivos a Github

        