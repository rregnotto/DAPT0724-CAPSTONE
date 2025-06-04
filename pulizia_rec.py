import pandas as pd
import os
import numpy as np
import csv

path = "C:\\Users\\franc\\OneDrive\\Desktop\\Workspace\\capstone"
file = os.path.join(path, "reclami.xlsx")

#Verifico il path del file
print("Path file:", file)
print(os.path.exists(file))

#Leggo il file excel 
df = pd.read_excel(file, engine="openpyxl", header=None)
print(df.head())
print("Coordinate da excel:")
print(df.iloc[:,29].head(10))  # Stampo le prime 5 righe della colonna Coordinate

#Stampando le prime righe della colonna Coordinate, vedo che ho assegnato il nome "Coordinate" alla colonna sbagliata. Lo cambio manualmente.

#Rinomino le colonne
df.columns = ["RecID", "Cod_prod", "Prodotto", "Data_prod", "Mix", "X", "Stabilimento", 
              "X", "X", "Danno", "Data_fatt", "Fattura", "Data_rec", "Cod_cliente",
               "Stamp", "Data_stamp", "X", "X", "Provincia", "Regione", "Posa", "Area", 
               "X", "X", "Linea_Gronda", "X", "X", "Altitudine", "X", "X", "Coordinate", "X", "X", 
               "X", "X", "X"]
print(df.columns)

#Converto e salvo il file excel in csv
csv_path = os.path.join(path, "reclami.csv")
df.to_csv(csv_path, index=False, sep=";", quoting=csv.QUOTE_ALL) #csv.QUOTE_ALL per gestire il problemma della colonna coordinate
print(df.shape)

#Rimuovo le colonne non necessarie
df = df.drop(columns = ["X", "Data_stamp"])
print(df.columns)

#Converto in formato date le colonne con data
date_col = ["Data_prod", "Data_fatt", "Data_rec"]

for col in date_col:
    df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

print(df[date_col].head())

#Rimuovo gli spazi all'inizio e alla fine delle stringhe
df = df.apply(lambda x: x.str.strip() if x.dtype == "str" else x)

#Sistemazione delle colonne
#Mix
df["Mix"] = df["Mix"].str.replace("/", "")
df["Mix"] = df["Mix"].str.replace(" ", "", regex=False)
df["Mix"] = df["Mix"].str.replace("0", "ND", regex=False)
print(df["Mix"])

#ID
df["RecID"] = df["RecID"].str.replace("_", "", regex=False)
df["RecID"] = df["RecID"].str.replace(" ", "", regex=False)

#Prodotto
df["Prodotto"] = df["Prodotto"].str.strip()
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
df["Cod_cliente"] = df["Cod_cliente"].str.replace("00", "ND", regex=False)

#Stabilimento
df["Stabilimento"] = df["Stabilimento"].str.strip()
#Posa
df["Posa"] = df["Posa"].str.strip()
df["Posa"] = df["Posa"].str.lower()
df["Posa"] = df["Posa"].str.replace(" ", "_", regex=True)  # Case insensitive replacement
df["Posa"] = df["Posa"].str.replace("0", "ND")

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

#Linea_Gronda
df["Linea_Gronda"] = df["Linea_Gronda"].str.strip()
df["Linea_Gronda"] = df["Linea_Gronda"].str.lower()
df["Linea_Gronda"] = df["Linea_Gronda"].str.replace("0", "ND", regex=False)

#Coordinate: la virgola che separa latitudine e longitudine viene sostituita perché crea conflitto con il separatore del file csv
df["Coordinate"] = df["Coordinate"].astype(str) 

#Funzione per estrarre latitudine e longitudine dalla colonna "Coordinate"
def extract_coordinates(val):
    try:
        coord = [float(x.strip()) for x in val.split(",") if x.strip() != ""] #divido la str val in base alla , e rimuovo gli spazi
        if len(coord) >= 2:
            return coord[0], coord[1]
        else:
            return None, None
    except:
        return None, None

# Applica la funzione riga per riga
df[["latitudine", "longitudine"]] = df["Coordinate"].apply(lambda x: pd.Series(extract_coordinates(x)))

# Controllo
print(df[["Coordinate"]].head())

#Sostituzione dei valori nulli o 0 con ND

df = df.replace([np.nan, 0, 0.0, "0", "0.0"], "ND")

#Salvo il file csv pulito
file_csv = "rec_clean.csv"
rec_clean = os.path.join(path, file_csv)
df.to_csv(rec_clean, index=False)
print("File pulito salvato in:", rec_clean)