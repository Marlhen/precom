### Automatizaci[o]n con Selenium TSICAM
### By Marlhen Estrada Dic 2022 Rev 0.2
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

#### Ingresamos url, usuario y la contrasena.

username = "mestradav@southernperu.com.pe"
password = "Nuevaera202112"
url = "https://tsicam.southernperu.com.pe/intranet/documentos/area/revision"

#### Arrancamos lo automatizado

driver=webdriver.Chrome(executable_path="C:\\driver\\chromedriver.exe")
driver.get(url)

##Maximiza la ventana

#driver.maximize_window()

driver.find_element_by_name("username").send_keys(username)
driver.find_element_by_name("password").send_keys(password)
driver.find_element_by_css_selector("input.btn.btn-primary").click()
time.sleep(3)

#### Cargamos el excel donde estan los DNI[s]

filesheet = "C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\excel.xlsx"
wb = load_workbook(filesheet)
hojas = wb.get_sheet_names()
print(hojas)
nombres = wb.get_sheet_by_name('Sheet1')
wb.close()

#CREANDO LISTA DE LOS DNI´s a activar"
activados = [] ## para los que activamos
estados = [] ## para colocar los estados
yaactivados = [] ## para los que ya estaban activados
#### Hacemos For para la cantidad de ingresos
# dni, nomb = A y B son las columnas  y [0] empieza en la primera columna

for i in range(1,45):
    dni, nomb = nombres[f'A{i}:B{i}'][0]
    print(dni.value, nomb.value)
    time.sleep(3)
    driver.find_element_by_name("nombre").send_keys(dni.value)
    driver.find_element_by_name("nombre").send_keys(Keys.ENTER)
    time.sleep(3)

    try:
    ## Verifica si el selector existe.
        item = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,"//a[@title='Reagendar Induccion y Orientacion Básica de Seguridad - Cuajone']")))
   
    ## Si existe el item lo activa en el sistema
        driver.find_element_by_xpath("//a[@title='Reagendar Induccion y Orientacion Básica de Seguridad - Cuajone']").click()
        time.sleep(3)
        driver.find_element_by_name("fechaSolicitud").clear()
        time.sleep(3)
        driver.find_element_by_name("fechaSolicitud").send_keys("21/05/2022")
        time.sleep(3)
        driver.find_element_by_xpath("(//a[@id='save-modal'])[1]").click()
        time.sleep(3)
        driver.find_element_by_css_selector(".fa.fa-check").click()
        time.sleep(3)
        #colocar la nota
        driver.find_element_by_xpath("(//input[@id='id_calificacion'])[1]").send_keys("17")
        time.sleep(3)
        #Sube el anexo 4
        driver.find_element_by_css_selector("#id_anexo").send_keys("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\"+ str(dni.value) +".pdf")
        time.sleep(3)
        driver.find_element_by_css_selector("#save-modal").click()
        time.sleep(5)

        #Para hacer prueba y cancelar"
        #driver.find_element_by_xpath("(//a[@class='modal-close waves-effect waves-green btn-flat'][normalize-space()='Cancelar'])[1]").click()
        #time.sleep(1) 
        driver.find_element_by_name("nombre").clear()
        time.sleep(7)

        # Exportamos la lista de los activados

        activados.append(dni.value)
        estados.append("activados")
        data1 = {"DNI":activados,"Estado":estados}
        df1 = pd.DataFrame(data1)
        df1.to_excel ("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\numero_trabajadores_activados.xlsx", index = False, header=True)

    except TimeoutException as e:

        # Exportamos la lista de los que estaban activados
        yaactivados.append(dni.value)
        print("Ya estaba activado")
        estados.append("Ya estaban activados")
        data2 = {"DNI":yaactivados,"Estado":estados}
        df2 = pd.DataFrame(data2)
        df2.to_excel ("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\numero_trabajadores_estaban_activados.xlsx", index = False, header=True)

        pass
        time.sleep(3)

    driver.find_element_by_name("nombre").clear()
    time.sleep(3)