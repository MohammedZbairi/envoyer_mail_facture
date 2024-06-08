import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

# Email configuration
smtp_server = 'smtp.gmail.com'
smtp_port = 465
smtp_user = 'zmo.salamarket@gmail.com'
smtp_password = 'rpmdafixalzbkfxx'
to_address = f"mohamed.zbairi@salamarket31.fr"

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
        print(f"Failed to send test email: Connection unexpectedly closed - {e}")
    except Exception as e:
        print(f"Failed to send test email: {e}")


# Load the Excel file
file_path = 'C:/Users/admin/Desktop/mohammed_z/DEV/PYTHON/BL/20240605-162156_ZeenDoc_Indexes.xlsx'
try:
    df = pd.read_excel(file_path, header=3)  # Reading the file starting at the 4th row
    print("Excel file loaded successfully.")
except FileNotFoundError:
    print(f"Error: The file at {file_path} was not found.")
    exit(1)
except Exception as e:
    print(f"An error occurred while reading the Excel file: {e}")
    exit(1)

# Loop through each row in the dataframe
for index, row in df.iterrows():
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
        subject = f"Demande de facture"
        body = f"""

        Bonjour,

        Veuillez nous transmettre les factures des Bons de Livraison suivant ( et ci-joint ):

        - BL {reference_document} de la livrason du {date_document}.

        Cordialement,
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

