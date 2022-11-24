import smtplib
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import string
import random



def random_generator(size=6, chars=string.ascii_uppercase + string.digits):
 return ''.join(random.choice(chars) for _ in range(size))

def send_email(emails, content, subject):
    try:  
        corpo_email = content

        msg = email.message.Message()
        msg['Subject'] = subject
        msg['From'] = 'agenciamktm@gmail.com'
        msg['To'] = emails
        password = 'vqslnaabrctovbbm' 
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email )

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('Email enviado')
    except Exception as e:
        print(f"Erro ao enviat email: {e}") 


def send_email_with_anexo(email, content, subject, dir):
    try:
        smtp_ssl_host = "smtp.gmail.com"
        smtp_ssl_port = 465

        recipient = email
        username = "agenciamktm@gmail.com"
        password = 'vqslnaabrctovbbm'

        message = MIMEMultipart()
        message["From"] = username
        message["To"] = ", ".join(recipient)
        message["Subject"] = subject
        body = MIMEText(content)
        message.attach(body)

        fp = open(dir, 'rb')
        anexo = MIMEApplication(fp.read(), _subtype="pptx")
        fp.close()
        anexo.add_header('Content-Disposition', 'attachment', filename='VSL')
        message.attach(anexo)

        server = smtplib.SMTP_SSL(smtp_ssl_host, smtp_ssl_port)
        server.login(username, password)
        server.sendmail(username, recipient, message.as_string().encode('utf-8'))
        server.quit()
    except Exception as e:
        print(f"Erro ao enviat email: {e}") 