import os, shutil

for root, dirs, files in os.walk("C:\\Users\\REDMIBOOK 16\\Downloads\\OneDrive_2022-05-21\\"):
    for file in files:
        if file[-4:].lower() == '.pdf':
            shutil.copy(os.path.join(root, file), os.path.join("C:\\Users\\REDMIBOOK 16\\Downloads\\prueba\\", file))