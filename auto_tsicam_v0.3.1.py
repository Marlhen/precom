### Automatizaci[o]n con Selenium TSICAM
### By Marlhen Estrada Dic 2022 Rev 0.3.3
### Fecha: 10.06.2022

from xmlrpc.client import boolean
from numpy import busday_count, imag
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from win10toast import ToastNotifier
import pandas as pd
import os
import time

######  
##from Screenshot import Screenshot_Clipping
from PIL import Image

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
poradministracion = [] ## para los que no aparecen en el sistema

### For en el rango  0 hasta el numero de filas
for i in range(0, int(numerofilas)):
    buscarpordni = dnis["dni"][i] ### buscar por dni igual a la fila del Data frame DNI
    print(buscarpordni)

    driver.find_element(by=By.NAME, value="nombre").send_keys(str(buscarpordni))
    time.sleep(3)
    driver.find_element(by=By.NAME, value="nombre").send_keys(Keys.ENTER)
    time.sleep(3)
    
    ## Verifica si el selector existe.

    at = driver.find_elements(by=By.XPATH, value="(//a[@title='Reagendar Induccion y Orientacion Básica de Seguridad - Cuajone'])[1]")

    ###------------------Devuelve si es falso o verdadero el selector buscado
    print(boolean(at))
    
    if len(at) > 0:
        print ("Voy a activarlos")
        ## Si existe el item lo activa en el sistema
        driver.find_element(by=By.XPATH, value="(//a[@title='Reagendar Induccion y Orientacion Básica de Seguridad - Cuajone'])[1]").click()
        time.sleep(3)
        driver.find_element(by=By.NAME, value="fechaSolicitud").clear()
        time.sleep(3)
        driver.find_element(by=By.NAME, value="fechaSolicitud").send_keys("26/05/2022")
        time.sleep(3)
        driver.find_element(by=By.XPATH, value="(//a[@id='save-modal'])[1]").click()
        time.sleep(3)
        driver.find_element(by=By.XPATH, value="(//i[@class='fa fa-check'])[1]").click()
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
          
            print("No esta para aprobar")
            ##### Notificacion dentro del panel
            toaster = ToastNotifier()
            toaster.show_toast("TSICAM", str(buscarpordni) + " " + "No esta para aprobar", duration=6)
        
            ##########

            
            #####-------------Buscar en que estado el pase...............##########

            driver.get("https://tsicam.southernperu.com.pe/intranet/pases/")
            time.sleep(2)
            driver.find_element(by=By.NAME, value="nombre").send_keys(str(buscarpordni))
            time.sleep(2)
            driver.find_element(by=By.NAME, value="nombre").send_keys(Keys.ENTER)
            time.sleep(2)

            atxx = driver.find_elements(by=By.XPATH, value="(//a[normalize-space()='100 %'])[1]")

            if len(atxx) > 0:
                driver.find_element(by=By.XPATH, value="(//a[normalize-space()='100 %'])[1]").click()
                time.sleep(2)
                
                atyy = driver.find_elements(by=By.XPATH, value="//td[contains(text(),'Esperando aprobación de')]")
                
                if len(atyy) > 0:
                    driver.find_element(by=By.XPATH, value="//td[contains(text(),'Esperando aprobación de')]").click()
        
                    #####-------------Captura la pantalla...............##########+

                    ###os.mkdir("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\imagenes\\")
                    driver.save_screenshot("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\imagenes\\" + str(buscarpordni) +"_" + "Myimage.png")
                    driver.get("https://tsicam.southernperu.com.pe/intranet/documentos/area/revision")
                    time.sleep(2)
                    
                    print("Falta aprobar por administracion")
                    ##### Notificacion dentro del panel
                    toaster = ToastNotifier()
                    toaster.show_toast("TSICAM", str(buscarpordni) + " " + "Falta aprobar por administracion", duration=6)

                    poradministracion.append(buscarpordni)
                    dfcsc2 = pd.DataFrame(poradministracion)
                    dfcsc2.columns = ["Por administracion"]
                    dfcsc2.to_csv("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\por_administracion1.csv", index = False, header=True)
                    
                else:

                    print("Falta aprobar por administracion")
                    ##### Notificacion dentro del panel
                    toaster = ToastNotifier()
                    toaster.show_toast("TSICAM", str(buscarpordni) + " " + "Falta aprobar por administracion", duration=6)
                    
                    #####-------------Captura la pantalla...............##########+
                    ##os.mkdir("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\imagenes1\\")
                    driver.save_screenshot("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\imagenes1\\" + str(buscarpordni) +"_" + "Myimage.png")
                    driver.get("https://tsicam.southernperu.com.pe/intranet/documentos/area/revision")
                    time.sleep(2)

                    poradministracion.append(buscarpordni)
                    dfcsc2 = pd.DataFrame(poradministracion)
                    dfcsc2.columns = ["Por administracion"]
                    dfcsc2.to_csv("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\por_administracion.csv", index = False, header=True)
                    
                    driver.get("https://tsicam.southernperu.com.pe/intranet/documentos/area/revision")
                    time.sleep(2)
                    
            else:

                print("No esta en sistema")
                ##### Notificacion dentro del panel
                toaster = ToastNotifier()
                toaster.show_toast("TSICAM", str(buscarpordni) + " " + "No esta en sistema", duration=6)

                driver.get("https://tsicam.southernperu.com.pe/intranet/documentos/area/revision")
                time.sleep(2)
                
                nosistema.append(buscarpordni)
                dfcsc3 = pd.DataFrame(nosistema)
                dfcsc3.columns = ["No sistema"]
                dfcsc3.to_csv("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\no_sistema.csv", index = False, header=True)
            
    
    driver.find_element(by=By.NAME, value="nombre").clear()
    time.sleep(3)




    