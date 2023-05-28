import smtplib
import ssl
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time

# Read excel file
df = pd.read_excel('emails.xlsx')

# Read body text from file
with open('body.txt', 'r') as f:
    body_text = f.read()

# Iterate over emails
for email in df['Emails']:
    print("Sending email to {}".format(email))

    msg = MIMEMultipart()
    msg['From'] = 'email@hotmail.com'  # Change this to your Outlook email
    msg['To'] = email
    msg['Subject'] = 'The subject of your email'

    msg.attach(MIMEText(body_text, 'plain'))

    # Open file in binary mode
    with open('Resume.pdf', 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    part.add_header(
        'Content-Disposition',
        f'attachment; filename= "Resume.pdf"',
    )

    msg.attach(part)
    text = msg.as_string()

    # Connect to the server and send the email
    server = smtplib.SMTP('smtp-mail.outlook.com', 587, timeout=10)
    server.starttls()
    server.login("email@hotmail.com", "password")  # Change these to your Outlook email and password
    server.sendmail('email@hotmail.com', email, text)  # Change 'email@hotmail.com' to your Outlook email
    server.quit()

    print("email sent to {}".format(email))
    time.sleep(5)  # Delay for 5 seconds
