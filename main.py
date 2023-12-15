import pyodbc
import pandas as pd
import SQL as s
from datetime import datetime
import time
from tkinter import messagebox




#start = time.time()

# Query Connection
connx_string = r"DRIVER={SQL Server}; server=172.19.128.2\emeadb; database=emea_enventa_live; UID=usr_razek; PWD=wB382^%H3INJ"
conx = pyodbc.connect(connx_string)

datum = str(datetime.now())[0:10]

sqls = [[s.artikel, 'Artikelstamm'], [s.artikeleinheit, 'Artikeleinheit'], [s.artikeltext, 'Artikelbeschreibung'], [s.produktentwicklung, 'Produktentwicklung'],
        [s.rabatt_kopf, 'Rabatt_Kopf'], [s.rabatt_gueltig, 'Rabatt_Gueltig'], [s.rabatt_staffel, 'Rabatt_Staffel'], [s.stueckliste, 'Stueckliste']]


for sql in sqls:
    try:
        df= pd.read_sql(sql[0], conx)
        df.to_excel(fr'S:\EMEA\Aenderungsprotokoll\{sql[1]}\{sql[1]}_{datum}.xlsx', index=False)
    except PermissionError:
        messagebox.showinfo('Excel offen!', f'Bitte schließe die Excel Liste: {sql[1]} und wiederhole den Vorgang.')
        break
    except OSError:
        messagebox.showinfo('Ordner Gelöscht', f'Kontrolliere ob der Ordner: {sql[1]} vorhanden ist und wiederhole den Vorgang.')
        break

#end = time.time()
print('Backup abgeschlossen')

messagebox.showinfo('Enventa Backup Erfolgreich!', 'Es wurde erfolgreich ein Backup angelegt.')

