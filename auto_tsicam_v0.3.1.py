### Automatizaci[o]n con Selenium TSICAM
### By Marlhen Estrada Dic 2022 Rev 0.3.1
from xmlrpc.client import boolean
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from win10toast import ToastNotifier
import pandas as pd
import time

#### Ingresamos url, usuario y la contrasena.

username = "mestradav@southernperu.com.pe"
password = "Nuevaera202112"
url = "https://tsicam.southernperu.com.pe/intranet/documentos/area/revision"

#### Arrancamos automatico el navegador Chrome y mandamos URL

navegador=Service("C:\\driver\\chromedriver.exe")
driver = webdriver.Chrome(service=navegador)
driver.get(url)

##Maximiza la ventana

#driver.maximize_window()
driver.find_element(by=By.NAME, value="username").send_keys(username)
driver.find_element(by=By.NAME, value="password").send_keys(password)
driver.find_element(by=By.CSS_SELECTOR, value="input.btn.btn-primary").click()
time.sleep(3)

#### Cargamos el excel donde estan los DNI[s]

filesheet = "C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\excel.xlsx"
### Cargamos todo el data frame con pd a dnis (tenemos un data frame)
dnis = pd.read_excel(filesheet, sheet_name="Sheet1")
### Seleccionamos la primera columna del data frame dnis a travez de columns dnis.columns[0:1] puede ser sucecivo 0:2 para mas columnas
selecolumdni = dnis[dnis.columns[0:1]]
### Contamos el numero de filas para hacer el for
numerofilas = selecolumdni.count(axis=0)

aprobados = [] ## lista para guardar los que vamos a aprobados
yaaprobados = [] ## lista para guardar los ya aprobadeos
nosistema = [] ## para los que no aparecen en el sistema

### For en el rango  0 hasta el numero de filas
for i in range(0, int(numerofilas)):
    buscarpordni = dnis["dni"][i] ### buscar por dni igual a la fila del Data frame DNI
    print(buscarpordni)

    driver.find_element(by=By.NAME, value="nombre").send_keys(buscarpordni)

    driver.find_element(by=By.NAME, value="nombre").send_keys(Keys.ENTER)
    time.sleep(3)
    
    ## Verifica si el selector existe.

    at = driver.find_elements(by=By.XPATH, value="(//a[@title='Reagendar Induccion y Orientacion Básica de Seguridad - Toquepala'])[1]")

    ###------------------Devuelve si es falso o verdadero el selector buscado
    print(boolean(at))
    
    if len(at) > 0:
        print ("Voy a activarlos")
        ## Si existe el item lo activa en el sistema
        driver.find_element(by=By.XPATH, value="(//a[@title='Reagendar Induccion y Orientacion Básica de Seguridad - Toquepala'])[1]").click()
        time.sleep(3)
        driver.find_element(by=By.NAME, value="fechaSolicitud").clear()
        time.sleep(3)
        driver.find_element(by=By.NAME, value="fechaSolicitud").send_keys("26/05/2022")
        time.sleep(3)
        driver.find_element(by=By.XPATH, value="(//a[@id='save-modal'])[1]").click()
        time.sleep(3)
        driver.find_element(by=By.XPATH, value="//i[@class='fa fa-check']").click()
        time.sleep(3)
        #colocar la nota
        driver.find_element(by=By.NAME, value="calificacion").send_keys("18")
        time.sleep(3)
        driver.find_element(by=By.NAME, value="archivo").send_keys("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\asistencia.pdf")
        time.sleep(3)
        #Sube el anexo 4
        driver.find_element(by=By.CSS_SELECTOR, value="#id_anexo").send_keys("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\"+ str(buscarpordni) +".pdf")
        time.sleep(3)
        driver.find_element(by=By.CSS_SELECTOR, value="#save-modal").click()
        time.sleep(5)

        #Para hacer prueba y cancelar"
        #driver.find_element_by_xpath("(//a[@class='modal-close waves-effect waves-green btn-flat'][normalize-space()='Cancelar'])[1]").click()
        #time.sleep(1) 
        driver.find_element(by=By.NAME, value="nombre").clear()
        time.sleep(3)

         ##### Notificacion dentro del panel
        toaster = ToastNotifier()
        toaster.show_toast("TSICAM", str(buscarpordni) + " " + "Se activo", duration=6)
            ##########

        # Exportamos la lista de los activados
        ##activados.append(dni.value)
        aprobados.append(buscarpordni)
        dfcsc = pd.DataFrame(aprobados)
        dfcsc.columns = ["Activados"]
        dfcsc.to_csv("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\activados.csv", index = False, header=True)

    else:
        at1 = driver.find_elements(by=By.XPATH, value="//span[@class='status green-text text-darken-2']")
        print(at1)
        
        if len(at1) > 0:
            print("Esta aprobado")
            ##### Notificacion dentro del panel
            toaster = ToastNotifier()
            toaster.show_toast("TSICAM", str(buscarpordni) + " " + "Ya estaba aprobado", duration=6)
           
            ##########
            yaaprobados.append(buscarpordni)
            dfcsc1 = pd.DataFrame(yaaprobados)
            dfcsc1.columns = ["Ya Activados"]
            dfcsc1.to_csv("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\ya_activados.csv", index = False, header=True)
            
        else:
          
            print("No esta en el sistema")
            ##### Notificacion dentro del panel
            toaster = ToastNotifier()
            toaster.show_toast("TSICAM", str(buscarpordni) + " " + "No esta en el sistema", duration=6)
            ##########
            nosistema.append(buscarpordni)
            dfcsc2 = pd.DataFrame(nosistema)
            dfcsc2.columns = ["No sistema"]
            dfcsc2.to_csv("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\no_sistema.csv", index = False, header=True)
    
    driver.find_element(by=By.NAME, value="nombre").clear()
    time.sleep(3)


    