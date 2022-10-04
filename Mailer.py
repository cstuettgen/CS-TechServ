import pathlib
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os
import csv


class Mailer:
    def __init__(self, first_name, last_name, to_email, carbon_copy, from_email_address,
                 smtp_password):
        self.first_name = first_name
        self.last_name = last_name
        self.to_email = to_email
        self.carbon_copy = carbon_copy
        self.subject = None
        self.msg = None
        self.from_email_address = from_email_address
        self.smtp_password = smtp_password
        self.compose_mail()
        self.embed_pics()
        self.attachments()

    def compose_mail(self):
        msg = MIMEMultipart('related')
        # msg['Subject'] = self.subject
        msg['From'] = self.from_email_address
        msg['To'] = self.to_email
        msg['cc'] = self.carbon_copy
        print(f"\nTO: {self.to_email}\n"
              f"FROM: {self.from_email_address}\n"
              f"CC: {self.carbon_copy}\n"
              )

        path = "Email_Message/email.html"

        with open(path, encoding='utf-8') as f:
            body_txt = f.read()
            formatted = body_txt
            print(f'Reading HTML file, "{path}", into email body...')
            with open('Email_Message/merge_data.txt') as merge_data:
                for line in merge_data:
                    if not line.strip():
                        continue
                    stripped_line = line.strip()
                    line_list = stripped_line.split()
                    if line_list[0] == "*":
                        continue
                    if 'Subject:' in line_list:
                        self.subject = " ".join(line_list[1:None])
                    else:
                        formatted = formatted.replace(line_list[0], line_list[1])
                        print(f"BODY: replacing {line_list[0]}, with {line_list[1]}")
            with open('Email_Message/merge_data.txt') as merge_data:
                for line in merge_data:
                    if not line.strip():
                        continue
                    stripped_line = line.strip()
                    line_list = stripped_line.split()
                    if line_list[0] == "*":
                        continue
                    if 'Subject:' in line_list:
                        continue
                    else:
                        self.subject = self.subject.replace(line_list[0], line_list[1])
                        # print(f"SUBJECT: replacing {line_list[0]}, with {line_list[1]}")

            print('\nSUBJECT: ', self.subject)
            msg['Subject'] = self.subject
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
            print(f'\nEMBEDDED IMAGES: {file}')
            i = i + 1

    def attachments(self):
        attachment_dir = r'Attachments/'
        attachments = os.listdir(attachment_dir)

        for file in attachments:
            if os.path.isfile(f'{attachment_dir}{file}') is True:
                attachment = MIMEApplication(open(f'{attachment_dir}{file}', 'rb').read())
                attachment.add_header('Content-Disposition', 'attachment', filename=file)
                self.msg.attach(attachment)
                print(f'\nATTACHMENTS: {attachment.get_filename()}')

    def send_mail(self):
        with smtplib.SMTP('smtp.office365.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(self.from_email_address, self.smtp_password)
            smtp.send_message(self.msg)


def notify():

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

            email = Mailer(first_name, last_name, to_email, carbon_copy,
                           from_email_address, smtp_password)
            print(f'\nSending email to {first_name.upper().strip()} '
                  f'{last_name.upper().strip()} '
                  f'<{to_email.lower().strip()}>'
                  '\n---------------------------------'
                  )

            email.send_mail()


def main():
    notify()


if __name__ == '__main__':
    main()
