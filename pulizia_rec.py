import pandas as pd
import os
import numpy as np

path = "C:\\Users\\franc\\OneDrive\\Desktop\\Workspace\\capstone"
file = os.path.join(path, "reclami.xlsx")

#Verifico il path del file
print("Path file:", file)
print(os.path.exists(file))

#Leggo il file excel 
df = pd.read_excel(file, engine="openpyxl")
print(df.head())

#Converto e salvo il file excel in csv
csv_path = os.path.join(path, "reclami.csv")
df.to_csv(csv_path, index=False)
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
df = df.drop(columns = ["X", "Data_stamp"])
print(df.columns)

#Converto in formato date le colonne con data
date_col = ["Data_prod", "Data_Fatt", "Data_rec"]

for col in date_col:
    df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

print(df[date_col].head())

#Rimuovo gli spazi all'inizio e alla fine delle stringhe
df = df.apply(lambda x: x.str.strip() if x.dtype == "str" else x)

#Sostituisco i valori nulli con ND
df = df.fillna("ND")
print(df["Mix"]) #check se i valori sono stati sostituiti correttamente

#Sistemazione delle colonne
#Mix
df["Mix"] = df["Mix"].str.replace("/", "")
df["Mix"] = df["Mix"].str.replace(" ", "", regex=False)
df["Mix"] = df["Mix"].str.replace("0", "ND", regex=False)
print(df["Mix"])

#ID
df["ID"] = df["ID"].str.replace("_", "", regex=False)
df["ID"] = df["ID"].str.replace(" ", "", regex=False)

#Prodotto
df["Prodotto"] = df["Prodotto"].str.lower()
df["Prodotto"] = df["Prodotto"].str.replace(".", "", regex=False)
df["Prodotto"] = df["Prodotto"].str.replace(" ", "_")

#Danno
df["Danno"] = df["Danno"].str.strip()
df["Danno"] = df["Danno"].str.replace("(?)", "")
df["Danno"] = df["Danno"].str.replace(",", "", regex=False)

#Fattura
df["Fattura"] = df["Fattura"].str.strip()
df["Fattura"] = df["Fattura"].str.replace("_", "")
df["Fattura"] = df["Fattura"].str.replace(" ", "", regex=False)
#df["Fattura"] = df["Fattura"].str.replace("0", "ND", regex=False)

#Cod_cliente
df["Cod_cliente"] = df["Cod_cliente"].astype(str)
df["Cod_cliente"] = df["Cod_cliente"].str.replace(".", "", regex=False)

#Sostituisco gli spazi interni con _
df = df.apply(lambda x: x.str.replace(" ", "_", regex=False) if x.dtype == "str" else x)

#Stamp
df["Stamp"] = df["Stamp"].str.strip()
df["Stamp"] = df["Stamp"].str.replace(" ", "", regex=False)

#Provincia
df["Provincia"] = df["Provincia"].str.strip()
df["Provincia"] = df["Provincia"].str.replace("0", "", regex=False)
df["Provincia"] = df["Provincia"].str.replace(" ", "_", regex=False)

#Regione
df["Regione"] = df["Regione"].str.strip()
df["Regione"] = df["Regione"].str.replace("0", "", regex=False)
df["Regione"] = df["Regione"].str.replace(" ", "_", regex=False)

#Salvo il file csv pulito
file_csv = "rec_clean.csv"
rec_clean = os.path.join(path, file_csv)
df.to_csv(rec_clean, index=False)
print("File pulito salvato in:", rec_clean)