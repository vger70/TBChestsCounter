import argparse
import os.path
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import shutil
from datetime import datetime
import tblibrary

# Google Sheets scope
scope            = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Apps Script Scope
script_scope     = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/gmail.send"]

# start
parser = argparse.ArgumentParser()
parser.add_argument("--config", help="file di configurazione",required=True)
args = parser.parse_args()

config = tblibrary.ReadConfig(args.config)

#tblibrary.run_macro(config.secret, config.token, script_scope, config.script_id, "ImpostaDati")

if not os.path.exists(config.data_file):
    print("Nothing to do.")
    exit(1)

# Google Sheet Authentication
credentials = ServiceAccountCredentials.from_json_keyfile_name(config.credentials_file, scope)
client = gspread.authorize(credentials)
print("Authenticated on Google Sheet")

# Google Sheets reference
sheet = client.open_by_url(config.sheet_name).sheet1

try:
    # Leggi il file CSV in un DataFrame pandas
    df = pd.read_csv(config.data_file, header=None, encoding = "ISO-8859-1")

    # Converti il DataFrame in una lista di dizionari
    values = df.values.tolist()
    
    print("Data on csv parsed")
    
    try:
        # Aggiungi i nuovi dati in append al foglio di Google Sheets
        sheet.append_rows(values, value_input_option='USER_ENTERED')
        print("Import completed")
        
        # sposta file in archivio
        backupdata = config.data_file.replace(".txt", "_") + datetime.now().strftime("%Y%m%d_%H%M%S") + ".txt"
        shutil.move(config.data_file, backupdata)
        shutil.move(backupdata, config.archive_dir)
        print("CSV archived")
        
        tblibrary.run_macro(config.secret, config.token, script_scope, config.script_id, "ImpostaDati")

    except Exception as e:
        raise
except pd.errors.EmptyDataError:
    print("No data")
except IOError:
    print("I/O error")
    
exit(0)