import pandas as pd
import os

path = "/Users/francescaregnotto/Workspace/DAPT0724/CAPSTONE/"
file = os.path.join(path, "reclami.xlsx")

#Verifico il path del file
print("Path file:", file)
print(os.path.exists(file))

#Leggo il file excel 
df = pd.read_excel(file, engine="openpyxl")
print(df.head())

#Converto e salvo il file excel in csv
csv_path = os.path.join(path, "reclami.csv")
df.to_csv("reclami.csv", index=False)
df = pd.read_csv(csv_path)
print(df.head())

print(df.shape)

#Rinomino le colonne
df.columns = ["ID", "Codice", "Prodotto", "Data_prod", "Mix", "X", "Stabilimento", 
              "X", "X", "Danno", "Data_Fatt", "Fattura", "Data_rec", "Cod_cliente",
               "Stamp", "Data_stamp", "X", "X", "Provincia", "Regione", "Posa", "Area", 
               "X", "X", "Linea_Gronda", "X", "X", "slm", "Coordinate", "X", "X", "X", "X", 
               "X", "X", "X"]

#Rimuovo le colonne non necessarie
df = df.drop(columns = ["X"])
print(df.columns)

#Rimuovo 