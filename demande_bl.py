import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import time
import glob

# Email configuration
smtp_server = 'imap.planet-work.com'
smtp_port = 993
smtp_user = 'mohamed.zbairi@salamarket31.fr'
smtp_password = 'Mohamed!31*'
to_address = f"zmo.salamarket@gmail.com"

# Define a function to send emails
def send_email(to_address, subject, body):
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = to_address
    msg['Subject'] = subject

    try:
        msg.attach(MIMEText(body, 'plain'))


        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:  # Using SMTP_SSL for SSL/TLS
            server.set_debuglevel(1)  # Enable debug output
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, to_address, msg.as_string())
        print(f"Test email sent to {to_address}")
    except smtplib.SMTPServerDisconnected as e:
        print(f"Failed to send test email: Connection unexpectedly closed - ")
    except Exception as e:
        print(f"Failed to send test email: ")


# Load the Excel file
file_path_pattern  = './BL/*ZeenDoc_Indexes.xlsx'
files = glob.glob(file_path_pattern )
for file_path in files:
    try:
        df = pd.read_excel(file_path, header=3)  # Reading the file starting at the 4th row
        print("Excel file loaded successfully.")
        print(df.head())
    except FileNotFoundError:
        print(f"Error: The file at {file_path} was not found.")
        exit(1)
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        exit(1)

# Group documents by fournisseur
grouped = df.groupby('Fournisseur')

# Loop through each group (each fournisseur)
for fournisseur, group in grouped:
    # Initialize email body
    body = f"""
    Bonjour,

    Veuillez nous transmettre les bon de livraisons émargé correspondants aux factures suivantes (et ci-joint) :

    """
    
    # Add details for each document in the group
    # Loop through each row in the dataframe
    for index, row in group.iterrows():
        if pd.notna(row['Identifiant']):
        # Extract the necessary information
            identifiant = row['Identifiant']
            type_document = row['Type de document']
            fournisseur = row['Fournisseur']
            reference_document = row['Référence du document']
            date_document = row['Date du document']
            montant_ttc = row['Montant TTC']
                # Assume the email address can be derived from the fournisseur name (you need to adjust this part)
            to_address = f"mohamed.zbairi@salamarket31.fr"
            # Create the email subject and body
            subject = f"Demande de bon de livraison"
            body += f"""- BL {reference_document} de la livraison du {date_document}.\n"""

    body += """Cordialement,
    SalaMarket Toulouse

    Groupe K&A Food


    8 avenue Larrieu-Thibaud

    31100 TOULOUSE


    (+33) (0)5 34 56 44 25 
    """
    # Send the email
    try:
        send_email(to_address, subject, body)
        print(f"Email envoyé à {to_address}")
    except Exception as e:
        print(f"Échec de l'envoi de l'email à {to_address} : {e}")
    
    time.sleep(30)

