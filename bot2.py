import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os

smtp_server = 'smtp.office365.com'
port = 587

sender_email = 'ibrahima.traore75@hotmail.com'
app_password = 'sfsqypuklldkjtjx'
recipient_emails = ['mt39070@gmail.com', 'halimtra75@gmail.com']

# Créez le message
message = MIMEMultipart()
message['From'] = sender_email
message['To'] = ', '.join(recipient_emails)
message['Subject'] = 'Sujet de l\'e-mail'
body = 'Corps du message.'
message.attach(MIMEText(body, 'plain'))

# Première pièce jointe
filename1 = r'C:/Users/ibrah/OneDrive/Documents/Ibrahima/CV/CVphoto/CVPTRAOREIbrahima2023.pdf'
with open(filename1, 'rb') as file:
    part1 = MIMEApplication(file.read(), Name=os.path.basename(filename1))
    part1.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(filename1)}"')
    message.attach(part1)

# Deuxième pièce jointe
filename2 = r'C:/Users/ibrah/OneDrive/Documents/Ibrahima/CV/CVphoto/LettreMotivationCandSTRAOREIbrahima.pdf'
with open(filename2, 'rb') as file:
    part2 = MIMEApplication(file.read(), Name=os.path.basename(filename2))
    part2.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(filename2)}"')
    message.attach(part2)

# Établissez une connexion sécurisée avec le serveur SMTP
server = smtplib.SMTP(smtp_server, port)
server.starttls()

# Connectez-vous au compte Outlook
server.login(sender_email, app_password)

# Envoyez l'e-mail
server.sendmail(sender_email, recipient_emails, message.as_string())

# Fermez la connexion au serveur SMTP
server.quit()

print('E-mail avec deux pièces jointes envoyé avec succès !')
