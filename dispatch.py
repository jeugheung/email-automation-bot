from smtplib import SMTP_SSL, SMTP_SSL_PORT
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# MARK: - Smtp Variables
message = "Say Hello to new iPhone"
fromEmail = "akartsight@gmail.com"
toEmails = "naisen04@gmail.com"
emailSubject = "Apple WWDC 2021"

class Smtp:

    def __init__(self, email, password):
        self.email = email
        self.password = password

    def connectSmtp(self):
        self.smtpServer = SMTP_SSL('smtp.gmail.com', port=SMTP_SSL_PORT)
        self.smtpServer.set_debuglevel(1)
        self.smtpServer.login(self.email, self.password)
        print("Connected")
        return self.smtpServer

    def sendMail(self, fromEmail, toEmails, emailSubject, message):
        msg = MIMEMultipart()
        msg['From'] = fromEmail
        msg['To'] = toEmails
        msg['Subject'] = emailSubject
        msg.attach(MIMEText(message, 'plain'))
        self.smtpServer.sendmail(msg['From'], msg['To'], msg.as_string())
        self.smtpServer.quit()

# MARK: - Smtp Methods
sending = Smtp(email, password)