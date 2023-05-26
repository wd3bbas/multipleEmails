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

# Create a secure SSL context
context = ssl.create_default_context()

# Use smtplib to send the email
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login("email@gmail.com", "Password")

    # Iterate over emails
    for email in df['Emails']:
        print("Sending email to {}".format(email))
        msg = MIMEMultipart()

        msg['From'] = 'email@gmail.com'
        msg['To'] = email
        msg['Subject'] = 'The subject of your email'

        msg.attach(MIMEText(body_text, 'plain'))

        # Open file in bynary mode
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

        server.sendmail('mohed.alabbas@gmail.com', email, text)

        print("email sent to {}".format(email))
        time.sleep(5)