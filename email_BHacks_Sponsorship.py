import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import configparser
import docx
import os
import pandas as pd

# Your email credentials
config = configparser.ConfigParser()
config.read('config.ini')

# Sender email
email_address = config['EMAIL_CREDENTIALS']['EMAIL_ADDRESS']
app_password = config['EMAIL_CREDENTIALS']['EMAIL_PASSWORD']

# Email setup
smtp_server = "smtp.gmail.com"
port = 465

# Member email
member_email = config['MEMBER_EMAIL']['EMAIL'].split(', ')

# Load email list from CSV
csv_file = 'sponsorships2024.csv'  # Path to CSV file
df = pd.read_csv(csv_file)

#Verify columns exist in CSV
if not {'Email', 'za'}.issubset(df.columns):
    raise ValueError("CSV file must contain 'Email' and 'za' columns")


def make_email(company_name, sender_name, member_in_charge=None):
    doc_path = 'templates/software.docx'
    body = ''
    if os.path.exists(doc_path):
        doc = docx.Document(doc_path)
        for para in doc.paragraphs:
            if "[Subject]" in para.text:
                body += '\n' + para.text.replace('[Subject]', 'Sponsorship Inquiry')
            elif "[company name/person name]" in para.text:
                body += '\n' + para.text.replace("[company name/person name]", company_name)
            elif "[your name]" in para.text:
                body += '\n' + para.text.replace("[your name]", sender_name)
            else:
                body += '\n' + para.text

    else:   
        print('Not a valid template path!')
        return None
    
    
    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = "" # recipient email
    # msg['CC'] = member_in_charge 
    msg['BCC'] = ','.join(member_email)
    msg['Subject'] = "Sponsorship Inquiry (Automatic Email Testing)"
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
    for index, row in df.iterrows():
        recipient_email = row['Email']
        company_name = row['za']
        
        email = make_email(company_name, recipient_email)
        if email:
            send_email(email)

if __name__ == "__main__":
    main()


