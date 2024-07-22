import os
import smtplib, ssl
from email.message import EmailMessage
import csv
from itertools import islice
import docx
from email.mime.text import MIMEText


# email id and passwoord
EMAIL_ADDRESS =  "" #config("EMAIL_BU_USER")
EMAIL_PASSWORD = "" #config("EMAIL_BU_PASS")
THIS_WEEK = 3
OTHER = []

# emails to CC (person in charge)
ANDREW = "andrewhu@bu.edu"
ANNIE = "aaliang@bu.edu"
HOANG = "nxhoang@bu.edu"
JASMINE = "fanchuz@bu.edu"
LUCAS = "junsunyoon02@gmail.com"
SHRADDHA = "slulla@bu.edu"
SOPHIE = "soheelim@bu.edu"

def starter():
    # get email templates and save it
    email_body = {}
    doc = docx.Document('Sponsorships_Email.docx')
    name = False
    holder = None
    for para in doc.paragraphs:
        if para.text == '':
            pass
        if para.text == "*****":
            name = True
        elif name:
            email_body[para.text] = []
            holder = para.text
            name = False
        else:
            email_body[holder].append(para.text)
    return email_body

def set_email(msg):
    #set up email (add sponsorships package etc..)
    msg['Subject'] = "Hello from BostonHacks!"
    with open('Sponsorship Packet 2023.pdf', 'rb') as f:
        file_data = f.read()
        file_name = f.name
    msg.add_attachment(file_data, maintype='image', subtype='pdf', filename=file_name)

    
    return msg


def make_email(company_name, person_name, address, type, tt, member_charge):
    # make email
    #msg["Bcc"] = "junsunyoon02@gmail.com"
    msg = EmailMessage()
    msg = set_email(msg)
    msg['To'] = address
    bdy = ""

    if type == "Tech(Software)":
        one = '\n'.join(email_body["Tech(Software)"]).replace("[first name]", person_name)
        one = one.replace("[company name]", company_name)
            
    elif type == "Tech(Hardware)":
        one = '\n'.join(email_body["Tech(Hardware)"]).replace("[first name]", person_name)
        one = one.replace("[company name]", company_name)
         
    elif type == "Finance":
        one = '\n'.join(email_body["Finance"]).replace("[first name]", person_name)
        one = one.replace("[company name]", company_name)
    
    elif type == "Food":
        one = '\n'.join(email_body["Food"]).replace("[first name]", person_name)
        one = one.replace("[company name]", company_name)
    
    else:
        one = '\n'.join(email_body["Other"]).replace("[first name]", person_name)
        one = one.replace("[company name]", company_name)


    if member_charge.strip() == "Andrew Hu":
        msg["cc"] = ANDREW
        msg['From'] = "Andrew Hu <" + EMAIL_ADDRESS + ">"
        one = one.replace("[your name]", "Andrew Hu")
    
    elif member_charge.strip() == "Annie Liang":
        msg["cc"] = ANNIE
        msg['From'] = "Annie Liang <" + EMAIL_ADDRESS + ">"
        one = one.replace("[your name]", "Annie Liang")
    
    elif member_charge.strip() == "Hoang Nguyen":
        msg["cc"] = HOANG
        msg['From'] = "Hoang Nguyen <" + EMAIL_ADDRESS + ">"
        one = one.replace("[your name]", "Hoang Nguyen")
    
    elif member_charge.strip() == "Jasmine Fanchu Zhou":
        msg["cc"] = JASMINE
        msg['From'] = "Jasmine Fanchu Zhou <" + EMAIL_ADDRESS + ">"
        one = one.replace("[your name]", "Jasmine Fanchu Zhou")
    
    elif member_charge.strip() == "Shraddha Lulla":
        msg["cc"] = SHRADDHA
        msg['From'] = "Shraddha Lulla <" + EMAIL_ADDRESS + ">"
        one = one.replace("[your name]", "Shraddha Lulla")
    
    elif member_charge.strip() == "Sophie Lim":
        msg["cc"] = SOPHIE
        msg['From'] = "Sophie Lim <" + EMAIL_ADDRESS + ">"
        one = one.replace("[your name]", "Sophie Lim")
    
    else:
        #msg["cc"] = LUCAS
        #msg['From'] = "Lucas Yoon <" + EMAIL_ADDRESS + ">"
        one = one.replace("[your name]", "Junsun (Lucas) Yoon [Head of Sponsorship]")
   


    body = MIMEText(one, "plain")
    msg.attach(body)

    
    return msg





if __name__ == "__main__":   
    #read contacts
    email_body = starter()
    with open('./emailsList/S4.csv') as csvfile:
        spamreader = csv.reader(csvfile, quotechar='|')
        thisweek = False
        context=ssl.create_default_context()


        with smtplib.SMTP('smtp.gmail.com', port = 587) as smtp:
            smtp.starttls(context=context)

            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            for r in islice(spamreader, 1, None):
                if not thisweek:
                    if ("Week" in r[0] and r[0][-1:] == str(THIS_WEEK)):
                        thisweek = True
                        print("passing2")
                    pass
                else:
                    try:
                        print(r, "abcde")
                        if "Week" in r[0]:
                            pass
                        elif r[1] == "No":
                            mail = make_email(r[2], r[3], r[4], r[5], r[4], r[0])
                            smtp.send_message(mail)
                            # smtp.send()
                            print("email sent to", r[2], r[3])
                    except Exception as err:
                        print(err, "error message")
                        print("Done Sending!")
            





