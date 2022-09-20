import pathlib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os
import csv


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
        self.embed_pics()

    def compose_mail(self):
        msg = MIMEMultipart('related')
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

    def embed_pics(self):
        pic_dir = r'images/'
        pics = os.listdir(pic_dir)

        pic_list = []
        pic_ext = ['.jpg', '.jpeg', 'png']
        for file in pics:
            extension = pathlib.Path(file).suffix
            if extension.lower() in pic_ext:
                pic_list.append(file)

        pic_list.sort()
        i = 0

        for file in pic_list:
            with open(f'{pic_dir}{file}', 'rb') as img:
                image = MIMEImage(img.read())

            image.add_header('content-ID', f'<image{i}>')
            self.msg.attach(image)
            i = i + 1

    def send_mail(self):
        with smtplib.SMTP('smtp.office365.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(self.from_email_address, self.smtp_password)
            smtp.send_message(self.msg)


def main():

    from_email_address = None
    smtp_password = None

    with open('config/credentials.txt') as creds:
        for line in creds:
            if not line.strip():
                continue

            stripped_line = line.strip()
            line_list = stripped_line.split()

            if line_list[0] == "*":
                continue
            if line_list[0] == "Email:":
                from_email_address = " ".join(line_list[1:None])
            if line_list[0] == "Password:":
                smtp_password = ' '.join(line_list[1:None])

    with open('config/recipients.csv', 'r') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        _ = next(csv_reader)

        for row in csv_reader:
            first_name, last_name, to_email, carbon_copy = row
            subject = f'Hello {first_name}, this is a test message!'

            body_vars = (first_name, last_name, to_email)

            test = Mailer(first_name, last_name, to_email, subject, carbon_copy,
                          from_email_address, smtp_password, body_vars)
            test.send_mail()


if __name__ == '__main__':
    main()
