import sys
import threading
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.exceptions import RefreshError

def progress_bar(i, total, interval=0.1):
    x = int(100.0 * i / total)
    sys.stdout.write(f"\r[{'=' * x}{' ' * (100 - x)}] {100.0*i/total:.1f}%")
    sys.stdout.flush()
    
# ReadConfig -------------------------------------------------------------------------------------------
class KeyPress(threading.Thread):
    def run(self):
        self.input = input()
        
# ReadConfig -------------------------------------------------------------------------------------------
class ReadConfig(object):
    """ This class parses configuration parameters from the configuration file """
# ------------------------------------------------------------------------------------------------------
    def __init__(self, config_file):
        
        config_kvp = {}

        try:

            with open(config_file) as cfp:
                for cl, config_line in enumerate(cfp):
                
                    kvp_string = config_line.strip()
                    if len(kvp_string) == 0 or kvp_string[0] == "#": # kvp_string.find("#") > -1: # ignore empty lines and comments in config file
                        continue
                
                    kvp = kvp_string.split("=")
                    if len(kvp) != 2:
                        i = 0
                        k = kvp[1]
                        for x in kvp:
                            if i > 1:
                                k += "=" + x
                            i += 1
                        kvp[1] = k
                        
                    config_kvp[kvp[0].strip()] = kvp[1].strip()
        except:
            print("### FATAL ERROR: Unable to open or read the configuration file: {}").format(config_file)
            exit(-1)
        else:
            self.service_user        = config_kvp.get('service_user','').replace("\"","")
            self.credentials_file    = config_kvp.get('credentials_file','').replace("\"","")
            self.secret              = config_kvp.get('secret','').replace("\"","")
            self.token               = config_kvp.get('token','').replace("\"","")
            self.TBconfig_file       = config_kvp.get('TBconfig_file','').replace("\"","")
            self.data_file           = config_kvp.get('data_file','').replace("\"","")
            self.archive_dir         = config_kvp.get('archive_dir','').replace("\"","")
            self.sheet_name          = config_kvp.get('sheet_name','').replace("\"","")
            self.script_id           = config_kvp.get('script_id','').replace("\"","")
            self.MAX_RETRY_COUNT     = int(config_kvp.get('MAX_RETRY_COUNT',2))
            self.MAX_CHEST           = int(config_kvp.get('MAX_CHEST',20))
            self.mailTo              = config_kvp.get('mailTo','').replace("\"","")
            self.SUBJECT             = config_kvp.get('SUBJECT','').replace("\"","")
            self.BROWSER             = config_kvp.get('BROWSER','Chrome')
            
#-----------------------------------------------------------------------------------------------
def getOAuth2Token(SCOPES, SECRET, TOKEN_PATH):
    # Apps Script Authentication
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
        
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError:
                flow = InstalledAppFlow.from_client_secrets_file(SECRET, SCOPES)
                creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(SECRET, SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open(TOKEN_PATH, "w") as tk:
            tk.write(creds.to_json())

        print("OAuth2 Token created")
    return creds

#-----------------------------------------------------------------------------------------------
def run_macro(secret, token, script_scope, script_id, macro):
    creds = getOAuth2Token(script_scope, secret, token)

    # popola i campi
    service = build('script', 'v1', credentials=creds)
    
    # Crea la richiesta per l'API
    request = {
        'function': f'{macro}',
        'devMode': True
    }
    
    # Chiama la funzione Apps Script
    response = service.scripts().run(body=request, scriptId=script_id).execute()

    # Stampa un messaggio di successo
    print("Macro eseguita con successo sul foglio Google.")
#-----------------------------------------------------------------------------------------------