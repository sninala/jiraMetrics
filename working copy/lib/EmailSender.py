import os
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders


class EmailSender(object):

    def __init__(self, config, email_subject, email_body):
        self.config = config
        self.subject = email_subject
        self.email_content = email_body

    def send_email(self):
        from_address = self.config.get('EMAIL', 'FROM_USER')
        to_addresses_str = self.config.get('EMAIL', 'TO_USERS')
        to_addresses = to_addresses_str.split(",")
        smtp_server = self.config.get('EMAIL', 'SMTP_SERVER')
        smtp_port = self.config.get('EMAIL', 'SMTP_PORT')
        smtp_password = self.config.get('EMAIL', 'SMTP_PASSWORD')
        msg = MIMEMultipart()
        print "sening email to {}".format(to_addresses_str)
        msg['From'] = from_address
        msg['To'] = to_addresses_str
        msg['Subject'] = self.subject
        email_body = self.email_content
        msg.attach(MIMEText(email_body, 'plain'))
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(from_address, smtp_password)
        text = msg.as_string()
        server.sendmail(from_address, to_addresses, text)
        server.quit()

    def send_email_with_file_attachment(self, attachment_file):
        from_address = self.config.get('EMAIL', 'FROM_USER')
        to_addresses_str = self.config.get('EMAIL', 'TO_USERS')
        to_addresses = to_addresses_str.split(",")
        smtp_server = self.config.get('EMAIL', 'SMTP_SERVER')
        smtp_port = self.config.get('EMAIL', 'SMTP_PORT')
        smtp_password = self.config.get('EMAIL', 'SMTP_PASSWORD')
        msg = MIMEMultipart()
        print "sening email to {}".format(to_addresses_str)
        msg['From'] = from_address
        msg['To'] = to_addresses_str
        msg['Subject'] = self.subject
        email_body = self.email_content
        msg.attach(MIMEText(email_body, 'plain'))
        file_name = os.path.basename(attachment_file)
        attachment = open(attachment_file, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % file_name)
        msg.attach(part)
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(from_address, smtp_password)
        text = msg.as_string()
        server.sendmail(from_address, to_addresses, text)
        server.quit()

