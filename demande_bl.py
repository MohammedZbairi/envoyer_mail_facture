import pandas as pd 
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import time
import glob

# Configuration de l'e-mail
smtp_server = 'smtp.planet-work.com'
smtp_port = 587  # Port STARTTLS
smtp_user = 'mohamed.zbairi@salamarket31.fr'
smtp_password = 'Mohamed!31*'

# Fonction pour envoyer un e-mail
def send_email(to_address, subject, body):
    if not to_address or to_address.lower() == "non":
        print(f"Adresse e-mail invalide pour {to_address}, e-mail ignoré.")
        return
    
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = to_address
    msg['Subject'] = subject
    try:
        msg.attach(MIMEText(body, 'plain'))
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.ehlo()
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, to_address, msg.as_string())
        print(f"Email envoyé à {to_address}")
    except Exception as e:
        print(f"Échec de l'envoi de l'email à {to_address} : {e}")

# Charger le fichier des fournisseurs
try:
    fournisseurs_df = pd.read_excel('C:/Users/ZBAIRI/Desktop/DEV/envoyer_mail_facture/BL/fournisseurs.xlsx')
    print("Fichier fournisseurs chargé avec succès.")
except Exception as e:
    print(f"Erreur lors de la lecture du fichier des fournisseurs : {e}")
    exit(1)

if 'FOURNISSEUR' not in fournisseurs_df.columns or 'E-mail' not in fournisseurs_df.columns:
    print("Les colonnes 'FOURNISSEUR' ou 'E-mail' sont absentes du fichier des fournisseurs.")
    exit(1)

email_mapping = dict(zip(fournisseurs_df['FOURNISSEUR'], fournisseurs_df['E-mail']))

file_path_pattern = r'C:/Users/ZBAIRI/Desktop/DEV/envoyer_mail_facture/BL/*ZeenDoc_Indexes.xlsx'
files = glob.glob(file_path_pattern)
if not files:
    print("Aucun fichier de factures trouvé.")
    exit(1)

for file_path in files:
    try:
        df = pd.read_excel(file_path, header=0)
        print(df.head(10))
        print(f"Fichier chargé : {file_path}")
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier {file_path} : {e}")
        continue

    if 'Fournisseur' not in df.columns:
        print(f"Colonne 'Fournisseur' absente dans le fichier {file_path}")
        continue

    grouped = df.groupby('Fournisseur')
    for fournisseur, group in grouped:
        email = email_mapping.get(fournisseur, None)
        if not email:
            print(f"Aucune adresse e-mail trouvée pour le fournisseur : {fournisseur}")
            continue
        
        subject = f"Demande de Bons de livraison : {fournisseur}"
        body = f"Bonjour,\n\nVeuillez nous transmettre les bons de livraison emargés correspondants aux factures suivants sur l'adresse mail facture@salamarket31.fr :\n\n"
        for _, row in group.iterrows():
            date_document = pd.to_datetime(row['Date du document'])
            formatted_date = date_document.strftime("%d-%m-%Y")
            if pd.notna(row.get('Référence du document')) and formatted_date:
                body += f"- FACTURE N° {row['Référence du document']} .\n"
        body += "\n------\nSalaMarket Toulouse\n\nGroupe K&A Food\n\n8 avenue Larrieu-Thibaud\n\n31100 TOULOUSE\n\n(+33) (0)5 34 56 44 25"

        try:
            send_email(to_address=email, subject=subject, body=body)
        except Exception as e:
            print(f"Erreur lors de l'envoi de l'email à {email} : {e}")
        time.sleep(30)
