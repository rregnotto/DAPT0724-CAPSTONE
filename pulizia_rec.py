import pandas as pd
import os
import numpy as np

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
print(df.head())
#Rimuovo le colonne non necessarie
df = df.drop(columns = ["X"])
print(df.columns)

#Rimuovo gli spazi all'inizio e alla fine delle stringhe
df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

#Sostituisco i valori nulli con None
df = df.replace(["0"], np.nan)
print(df["Mix"]) #check se i valori sono stati sostituiti correttamente

#Sistemo i valori nella colonna Mix
df["Mix"] = df["Mix"].str.replace("/", "")
df["Mix"] = df["Mix"].str.replace(" ", "", regex=False)

print(df["Mix"])

#Sostituisco gli spazi interni con _
df = df.apply(lambda x: x.str.replace(" ", "_", regex=False) if x.dtype == "object" else x)

#Converto in formato date le colonne con data

date_col = ["Data_prod", "Data_Fatt", "Data_rec"]

for col in date_col:
    df[col] = pd.to_datetime(df[col], format= "%d/%m/%Y", errors ='coerce')