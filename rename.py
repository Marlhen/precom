from itertools import count
import pandas as pd
import os

folder = "C:\\Users\\REDMIBOOK 16\\Downloads\\OneDrive_1_5-15-2022\\R19.-29-04-22\\CERTIFICADOS\\"
count = 1

for nombrearchivo in os.listdir(folder):
    
    # Construct old file name
    source = folder + nombrearchivo
    # Adding the count to the new file name and extension
    destination = folder + nombrearchivo.split('-')[0] + ".pdf"
    # Renaming the file
    #os.rename(source, destination)
    activados = os.rename(source, destination)
    #Creamos un diccionario
    count += 1
    print(str(activados) + " " + "mas mas")

print('New Names are los siguientes')
# verify the result

# Redirecciona la lista de donde tomará los datos
res1 = os.listdir("C:\\Users\\REDMIBOOK 16\\Downloads\\OneDrive_1_5-15-2022\\R19.-29-04-22\\CERTIFICADOS\\")
# Corta el nombre para todos los de la lista y devuelve a travez del [0] solo el nombre y de acuerdo a cada numero.
res = [x.split('.')[0] for x in res1]
#creando la data
data = {'dni':res}
df = pd.DataFrame(data)
df.to_excel ("C:\\Users\\REDMIBOOK 16\\Downloads\\OneDrive_1_5-15-2022\\R19.-29-04-22\\CERTIFICADOS\\excel.xlsx", index = False, header=False)

print(res)