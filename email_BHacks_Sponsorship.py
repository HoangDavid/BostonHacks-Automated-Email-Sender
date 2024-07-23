import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import configparser

# Your email credentials
config = configparser.ConfigParser()
config.read('config.ini')

# Sender email
email_address = config['EMAIL_CREDENTIALS']['EMAIL_ADDRESS']
app_password = config['EMAIL_CREDENTIALS']['EMAIL_PASSWORD']

# Email setup
smtp_server = "smtp.gmail.com"
port = 465


def make_email():
    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = "" # recipient email
    msg['Subject'] = "Hey there"
    body = "The thing actually work"
    msg.attach(MIMEText(body, 'plain'))

    return msg

def send_email(msg):
    server = smtplib.SMTP_SSL(smtp_server, port)
    try:
        server.login(email_address, app_password)
        text = msg.as_string()
        server.sendmail(email_address, msg['To'], text)
        print("Email sent successfully!")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        server.quit()


def main():
    ...

if __name__ == "__main__":
    main()



