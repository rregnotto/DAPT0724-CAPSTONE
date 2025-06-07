import pandas as pd
import os
import numpy as np
import csv
import random
import datetime
from datetime import timedelta
from faker import Faker

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
df.to_csv(csv_path, index=False, sep=";", quoting=csv.QUOTE_ALL) #csv.QUOTE_ALL per gestire il problema della colonna coordinate
print(df.shape)

#Rimuovo le colonne non necessarie
df = df.drop(columns = ["X", "Area", "Stamp", "Data_stamp"])
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

#Data_fatt: sostituisco i valori "1970-01-01" con date casuali tra il 2004 e il 2024
mask = df["Data_fatt"] == datetime.date(1970, 1, 1)

# Sostituisco solo i valori selezionati con date casuali
fake = Faker("it_IT") 
# Genero date casuali tra il 2004 e il 2024 per le righe che soddisfano la maschera
df.loc[mask, "Data_fatt"] = [
    fake.date_between(start_date=datetime.date(2004, 1, 1), end_date=datetime.date(2024, 12, 31)) 
    for _ in range(mask.sum())
    ]
df["Data_fatt"] = pd.to_datetime(df["Data_fatt"]).dt.date

#Data_rec: sostituisco i valori "1970-01-01" con date casuali tra il 2004 e il 2024
mask = df["Data_rec"] == datetime.date(1970, 1, 1)

# Sostituisco solo i valori selezionati con date casuali
fake = Faker("it_IT") 
# Genero date casuali tra il 2004 e il 2024 per le righe che soddisfano la maschera
df.loc[mask, "Data_rec"] = [
    fake.date_between(start_date=datetime.date(2004, 1, 1), end_date=datetime.date(2024, 12, 31)) 
    for _ in range(mask.sum())
    ]
df["Data_rec"] = pd.to_datetime(df["Data_rec"]).dt.date

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
df[["Latitudine", "Longitudine"]] = df["Coordinate"].apply(lambda x: pd.Series(extract_coordinates(x)))

# Controllo
print(df[["Coordinate"]].head())

#Sostituzione dei valori nulli o 0 con ND

df = df.replace([np.nan, 0, 0.0, "0", "0.0"], "ND")

#Dato che i dati sono in maggiorparte fittizi, manipolo i dati di alcune colonne per migliorarne la consistenza

#Popolo tutte le righe della colonna Cod_cliente con un valore fittizio (uso la libreria Faker)
fake = Faker("it_IT") #uso la localizzazione italiana
df["Cod_cliente"] = [fake.unique.bothify(text = "CL####") for _ in range(len(df))]
print(df["Cod_cliente"].head())


#Popolo la colonna Danno con 3 tipologie di danno, in modo casuale
danno = ["Sfaldatura", "Rottura", "Delaminazione"]
df["Danno"] = np.random.choice(danno, size=len(df))

#Completo le celle vuote della colonna Mix con delle miscele aggiuntive
mix = ["V16", "V32AR", "V31", "V34CAM", "N59"]
mask = df["Mix"] == "ND" #uso una maschera booleana
df.loc[mask, "Mix"] = np.random.choice(mix, size=mask.sum()) #assegno un mix casuale alle celle della maschera

#Creo le nuove colonne per la dimensione cliente
#"contatto", "Impresa", "Telefono", "Email"
df["Contatto"] = [fake.name() for _ in range(len(df))]
df["Impresa"] = [fake.company() for _ in range(len(df))]
df["Telefono"] = [fake.phone_number() for _ in range(len(df))]
df["Email"] = [fake.email() for _ in range(len(df))]

##Creo le nuove colonne per la tabella reclami
#"Status", "Data_risoluzione"
status = ["In attesa", "In lavorazione", "Risolto", "Chiuso"]
df["Status"] = np.random.choice(status, size=len(df))
df["Data_ris"] = [
    fake.date_between(
        start_date=datetime.date(2004, 1, 1), 
        end_date=datetime.date(2024, 12, 31))
        for _ in range(len(df))
]
df["Data_ris"] = pd.to_datetime(df["Data_ris"]).dt.date


#Per rendere veritieri i dati, faccio in modo che Data_ris sia successiva a Data_rec
#Mi assicuro che per i reclami "In lavorazione" non ci sia una data di risoluzione

df["Data_rec"] = pd.to_datetime(df["Data_rec"], errors="coerce")

# Solo per reclami non "In lavorazione", assegna una data di risoluzione dopo Data_rec
df["Data_ris"] = df.apply(
    lambda row: row["Data_rec"] + timedelta(days=random.randint(7, 90)) # Aggiungo da 7 a 90 giorni a Data_rec
    if row["Status"] != "In lavorazione" and pd.notnull(row["Data_rec"]) 
    else pd.NaT,
    axis=1
)

#Faccio la stessa operazione per Data_fatt, in modo che sia sempre postumo a Data_prod
df["Data_prod"] = pd.to_datetime(df["Data_prod"], errors="coerce").dt.date

df["Data_fatt"] = df["Data_prod"].apply(
    lambda x: x + timedelta(days=random.randint(7, 90)) if pd.notnull(x) else "ND"
)

#Per ottimizzare il database e le seguenti relazioni tra tabelle, creo un ID univoco per i codici prodotto uguali ma prodotto in anni diversi
df["Data_prod"]= pd.to_datetime(df["Data_prod"])
#Estraggo l'anno di produzione
df["Anno_prod"] = df["Data_prod"].dt.year.astype('Int64')

#Combino cod_prod e anno_prod per creare un ID univoco
df['Cod_prod_uni'] = df['Cod_prod'] + "_" + df['Anno_prod'].astype(str)


#Salvo il file csv pulito e sistemato
file_csv = "rec_clean.csv"
rec_clean = os.path.join(path, file_csv)
df.to_csv(rec_clean, index=False)
print("File pulito salvato in:", rec_clean)
