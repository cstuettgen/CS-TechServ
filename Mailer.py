import pathlib
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os
import csv

"""*******************************************************************************************
Add Mailer.py to same directory as your code.

Copy 'def notify():' method to the code you are running.
    modify the subject and body-vars to match your subject 
    and order of variables in the body of of your HTML email body.

Create folders 'Config', 'Attachments', 'Email_Message', 'Images'.

Add HTML email body to 'Email_Message' and title it 'email.html'.

Images to be embedded will be labeled in alphabetical order starting with index0.
Add images to your HTML email body, 'email.html'
    Ex.| <img width=469 height=263  src="cid:image0">
       | <img width=469 height=263  src="cid:image1">

Create credentials.txt in 'Config' folder and label creds that will be sending the email 
with Email: and Password: on separate lines.
    Ex.| Email:  john.doe@wherever.net
       | Password:  ***********

Create 'recipients.csv' and place it in the 'Config' folder.
    Add this header to file: 'first_name,last_name,to_email,carbon_copy'.
    Begin adding recipients on line 2 of csv.

Add any files to be attached to the 'Attachments' folder.

Add 'notify()' to your processes.
**********************************************************************************************"""


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
        self.attachments()

    def compose_mail(self):
        msg = MIMEMultipart('related')
        msg['Subject'] = self.subject
        msg['From'] = self.from_email_address
        msg['To'] = self.to_email
        msg['cc'] = self.carbon_copy

        path = "Email_Message/email.html"

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

    def attachments(self):
        attachment_dir = r'Attachments/'
        attachments = os.listdir(attachment_dir)
        print(attachments)
        for file in attachments:
            print(file)
            if os.path.isfile(f'{attachment_dir}{file}') is True:
                attachment = MIMEApplication(open(f'{attachment_dir}{file}', 'rb').read())
                attachment.add_header('Content-Disposition', 'attachment', filename=file)
                self.msg.attach(attachment)
                print(attachment)

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

            # ---edit subject and add variables for email body in order of appearance in your HTML email message--
            subject = f'Hello {first_name}, this is a test message'
            body_vars = (first_name, last_name, to_email)
            # ----------------------------------------------------------------------------------------------------

            email = Mailer(first_name, last_name, to_email, subject, carbon_copy,
                           from_email_address, smtp_password, body_vars)

            email.send_mail()


def main():
    notify()


if __name__ == '__main__':
    main()
