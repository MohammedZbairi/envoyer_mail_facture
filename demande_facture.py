import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

# Load the Excel file
file_path = 'C:/Users/admin/Desktop/mohammed z/DEV/PYTHON/BL/20240605-162156_ZeenDoc_Indexes.xlsx'  # Update with the correct path to your Excel file
df = pd.read_excel(file_path, header=3)  # Reading the file starting at the 4th row

# Email configuration
smtp_server = 'imap.planet-work.com'
smtp_port = 993
smtp_user = 'reception@salamarket31.fr'
smtp_password = 'Reception!31*'

# Define a function to send emails
def send_email(to_address, subject, body):
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = to_address
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

# Loop through each row in the dataframe
for index, row in df.iterrows():
    # Extract the necessary information
    identifiant = row['Identifiant']
    type_document = row['Type de document']
    fournisseur = row['Fournisseur']
    reference_document = row['Référence du document']
    date_document = row['Date du document']
    montant_ttc = row['Montant TTC']
    
    # Assume the email address can be derived from the fournisseur name (you need to adjust this part)
    # to_address = f"{fournisseur.replace(' ', '').lower()}@example.com"
    to_address = f"mohamed.zbairi@salamarket31.fr"
    
    # Create the email subject and body
    subject = f"Document {identifiant} - {type_document}"
    body = f"""
    Bonjour {fournisseur},

    Veuillez trouver ci-dessous les détails du document :

    Identifiant: {identifiant}
    Type de document: {type_document}
    Référence du document: {reference_document}
    Date du document: {date_document}
    Montant TTC: {montant_ttc}

    Cordialement,
    Votre Service Comptabilité
    """
    
    # Send the email
    send_email(to_address, subject, body)

print("Emails sent successfully.")
