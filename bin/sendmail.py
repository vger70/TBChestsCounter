import argparse
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from googleapiclient.discovery import build
import tblibrary

# Se si modifica queste SCOPES, eliminare il file token.json.
scope_mail = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/gmail.send"]

def invia_email(FROM, TO, SUBJECT, BODY):
    creds = tblibrary.getOAuth2Token(scope_mail, config.secret, config.token)

    # Costruisce il servizio
    service = build('gmail', 'v1', credentials=creds)

    # Crea il messaggio
    messaggio = MIMEMultipart()
    messaggio['to'] = TO
    messaggio['from'] = FROM
    messaggio['subject'] = SUBJECT
    messaggio.attach(MIMEText(BODY, 'plain'))
    raw_messaggio = base64.urlsafe_b64encode(messaggio.as_bytes()).decode()

    # Invia il messaggio
    messaggio = service.users().messages().send(userId="me", body={'raw': raw_messaggio}).execute()

# Sendmail ------------------------------------------------------------------------------------------------------
parser = argparse.ArgumentParser()
parser.add_argument("--config", help="file di configurazione",required=True)
parser.add_argument("--to", help="il destinatario della mail",required=False,default="m.mosti@gmail.com")
parser.add_argument("--subject", help="l'oggetto del messaggio",required=True)
parser.add_argument("--body", help="Il corpo del messaggio",required=True)
args = parser.parse_args()

config = tblibrary.ReadConfig(args.config)

invia_email(config.service_user, args.to, args.subject, args.body)
