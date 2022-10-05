import pathlib
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os
import csv


class Mailer:
    def __init__(self, from_email_address, smtp_password, **kwargs):
        self.first_name = kwargs['first_name']
        self.last_name = kwargs['last_name']
        self.to_email = kwargs['to_email']
        self.carbon_copy = kwargs['carbon_copy']
        self.subject = kwargs['subject']
        self.directory = kwargs['directory']
        self.msg = None
        self.from_email_address = from_email_address
        self.smtp_password = smtp_password
        self.default_email_dir = 'Email_Message\\'
        self.default_atachments_dir = 'Attachments\\'
        self.default_images_dir = 'Images\\'
        self.set_dirs()
        self.compose_mail(**kwargs)
        self.embed_pics()
        self.attachments()

    def set_dirs(self):
        if (self.directory not in os.listdir(self.default_email_dir)
                and self.directory not in os.listdir(self.default_images_dir)
                and self.directory not in os.listdir(self.default_atachments_dir)):
            print(f'\n---Using default folders for email body, embedded images and attachments---')

        elif (self.directory not in os.listdir(self.default_email_dir)
                or self.directory not in os.listdir(self.default_images_dir)
                or self.directory not in os.listdir(self.default_atachments_dir)):
            print(f'\n*** A folder is missing for recipient {self.first_name} {self.last_name} <{self.to_email}> ***')
            exit()

        elif (self.directory in os.listdir(self.default_email_dir)
                and self.directory in os.listdir(self.default_images_dir)
                and self.directory in os.listdir(self.default_atachments_dir)):
            self.default_atachments_dir = os.path.join(self.default_atachments_dir, self.directory)
            self.default_images_dir = os.path.join(self.default_images_dir, self.directory)
            self.default_email_dir = os.path.join(self.default_email_dir, self.directory)
            print(f'\n---Using folder, "{os.path.join(self.default_email_dir,self.directory)}"'
                  f', for email body, embedded images and attachments---')

    def compose_mail(self, **kwargs):
        msg = MIMEMultipart('related')
        msg['From'] = self.from_email_address
        msg['To'] = self.to_email
        msg['cc'] = self.carbon_copy
        print(f"\nTO: {self.to_email}\n"
              f"FROM: {self.from_email_address}\n"
              f"CC: {self.carbon_copy}\n"
              )

        with open(os.path.join(self.default_email_dir, 'email.html'), encoding='utf-8') as f:
            body_txt = f.read()
            formatted = body_txt
            print(f'Reading HTML file, "{self.default_email_dir}\\email.html", into email body...')

            for item in kwargs:
                f_item = '{'+item+'}'

                if f_item in self.subject:
                    self.subject = self.subject.replace(f_item, kwargs[item])
                if f_item in formatted:
                    formatted = formatted.replace(f_item, kwargs[item])

            msg['Subject'] = self.subject
            msg.attach(MIMEText(formatted, 'html'))
            print(f'\nSUBJECT: {self.subject}')
            self.msg = msg

    def embed_pics(self):
        pic_dir = self.default_images_dir + '/'
        pics = os.listdir(pic_dir)
        print(pic_dir)
        pic_list = []
        pic_ext = ['.jpg', '.jpeg', '.png']
        for file in pics:
            extension = pathlib.Path(file).suffix.lower()

            if extension in pic_ext:
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
        attachment_dir = self.default_atachments_dir + '/'
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
        header = next(csv_reader)

        for row in csv_reader:
            i = 0
            csv_row = {}
            for column in header:
                csv_row.update({column.split()[0]: row[i]})
                i += 1

            email = Mailer(from_email_address, smtp_password, **csv_row)
            print(f'\nSending email to {csv_row["first_name"].upper().strip()} '
                  f'{csv_row["last_name"].upper().strip()} '
                  f'<{csv_row["to_email"].lower().strip()}>'
                  '\n---------------------------------'
                  )

            email.send_mail()


def main():
    notify()


if __name__ == '__main__':
    main()
