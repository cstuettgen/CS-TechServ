import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class Mailer:
    def __init__(self, first_name, last_name, to_email, subject, carbon_copy, from_email_address,
                 smtp_password, body_vars):
        self.first_name = first_name
        self.last_name = last_name
        self.to_email = to_email
        self.carbon_copy = carbon_copy
        self.subject = subject
        self.body_vars = body_vars
        self.msg = None
        self.from_email_address = from_email_address
        self.smtp_password = smtp_password
        self.compose_mail()

    def compose_mail(self):
        msg = MIMEMultipart('alternative')
        msg['Subject'] = self.subject
        msg['From'] = self.from_email_address
        msg['To'] = self.to_email
        msg['cc'] = self.carbon_copy

        path = "email_message/email.html"

        with open(path, encoding='utf-8') as f:
            if self.body_vars is not None:
                body_txt = f.read()
                formatted = body_txt % self.body_vars
            else:
                body_txt = f.read()
                formatted = body_txt

        f.close()

        msg.attach(MIMEText(formatted, 'html'))

        self.msg = msg

    def send_mail(self):
        with smtplib.SMTP('smtp.office365.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(self.from_email_address, self.smtp_password)
            smtp.send_message(self.msg)


def main():
    first_name = 'first'
    last_name = 'last'
    to_email = 'recipient'
    subject = 'Test'
    carbon_copy = ''
    from_email_address = 'sender'
    smtp_password = 'app password'
    body_vars = (first_name, last_name, to_email)
    test = Mailer(first_name, last_name, to_email, subject, carbon_copy, from_email_address, smtp_password, body_vars)
    test.send_mail()


if __name__ == '__main__':
    main()
