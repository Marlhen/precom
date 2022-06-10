from ast import Index
from itertools import count
import pandas as pd
import os

folder = "C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\"
count = 1

#Creamos la lista para guardar los datos
#Lista y estados 
lista = []
estados = []

#Haciendo un for para todos los archivos
for nombrearchivo in os.listdir(folder):
    
    # Construct old file name
    source = folder + nombrearchivo
    # Adding the count to the new file name and extension
    destination = folder + nombrearchivo.split('-')[0] + ".pdf"
    os.rename(source, destination)
    count += 1

    lista.append(nombrearchivo.split('-')[0])
    estados.append("Nombre cambiado")
    #print(lista)
    # Renaming the file
    #os.rename(source, destination)
    
    #Creamos un diccionario
    #print(lista)
    #print(str(activados) + " " + "mas mas")
data1 = {"xx":lista,"yy":estados}
df1 = pd.DataFrame(data1)
df1.columns = ["dni","Estado"]
df1.to_excel ("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\numero_trabajadores.xlsx", index = False, header= False)
    

#print('New Names are los siguientes mas y mas')
# verify the result

# Redirecciona la lista de donde tomará los datos
res1 = os.listdir("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\")
# Corta el nombre para todos los de la lista y devuelve a travez del [0] solo el nombre y de acuerdo a cada numero.
res = [x.split('.')[0] for x in res1]
#creando la data
data = {'dni':res}
df = pd.DataFrame(data)
df.to_excel ("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\excel.xlsx", index = False, header=True)

#print(res)

#Repositorio para cargar los archivos a Github
#https://github.com/Marlhen/automatizacion_TSICAM.git