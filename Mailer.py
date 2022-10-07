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
        self.default_attachments_dir = 'Attachments\\'
        self.default_images_dir = 'Images\\'
        self.set_dirs()
        self.compose_mail(**kwargs)
        self.embed_pics()
        self.attachments()
        # print(self.msg)

    def set_dirs(self):
        # if (self.directory not in os.listdir(self.default_email_dir)
        #         and self.directory not in os.listdir(self.default_images_dir)
        #         and self.directory not in os.listdir(self.default_attachments_dir)):

        if self.directory == '':
            print(f'\n---Using default folders for email body, embedded images and attachments---')

        elif (self.directory not in os.listdir(self.default_email_dir)
                or self.directory not in os.listdir(self.default_images_dir)
                or self.directory not in os.listdir(self.default_attachments_dir)):
            print(f'\n*** One or more folders is missing for recipient {self.first_name} {self.last_name}'
                  f' <{self.to_email}> ***')
            exit()

        elif (self.directory in os.listdir(self.default_email_dir)
                and self.directory in os.listdir(self.default_images_dir)
                and self.directory in os.listdir(self.default_attachments_dir)):
            self.default_attachments_dir = os.path.join(self.default_attachments_dir, self.directory)
            self.default_images_dir = os.path.join(self.default_images_dir, self.directory)
            self.default_email_dir = os.path.join(self.default_email_dir, self.directory)
            print(f'\n---Using folder, "{os.path.join(self.default_email_dir,self.directory)}"'
                  f', for email body, embedded images and attachments---')

    def compose_mail(self, **kwargs):
        msg = MIMEMultipart('related')
        msg['From'] = self.from_email_address
        msg['To'] = self.to_email
        msg['cc'] = self.carbon_copy

        if self.directory != '':
            email_html = f'{self.directory}.html'
        else:
            email_html = 'email.html'

        with open(os.path.join(self.default_email_dir, email_html), encoding='utf-8') as f:
            body_txt = f.read()
            formatted = body_txt
            print(f'\n*** Reading HTML file, "{self.default_email_dir}\\email.html", into email body... ***')
            print(f"\nTO: {self.to_email}\n"
                  f"FROM: {self.from_email_address}\n"
                  f"CC: {self.carbon_copy}\n"
                  )

            for item in kwargs:
                f_item = '{'+item+'}'

                if f_item in self.subject:
                    self.subject = self.subject.replace(f_item, kwargs[item])
                if f_item in formatted:
                    formatted = formatted.replace(f_item, kwargs[item])

            msg['Subject'] = self.subject
            msg.attach(MIMEText(formatted, 'html'))
            print(f'\n  SUBJECT: {self.subject}')
            self.msg = msg

    def sort_pics(self):
        pic_dir = self.default_images_dir + '\\'
        pics = os.listdir(pic_dir)

        pic_list = []
        pic_ext = ['.jpg', '.jpeg', '.png']
        for file in pics:
            extension = pathlib.Path(file).suffix.lower()

            if extension in pic_ext:
                pic_list.append(file)
        pic_list.sort()
        return pic_list

    def embed_pics(self):
        i = 0
        pic_list = self.sort_pics()
        for file in pic_list:
            with open(os.path.join(self.default_images_dir, file), 'rb') as img:
                image = MIMEImage(img.read())

            image.add_header('content-ID', f'<image{i}>')
            self.msg.attach(image)
            print(f'\n      EMBEDDED IMAGES: {file}')
            i = i + 1

    def attachments(self):
        attachment_dir = self.default_attachments_dir + '\\'
        attachments = os.listdir(attachment_dir)

        for file in attachments:
            if os.path.isfile(f'{attachment_dir}{file}') is True:
                attachment = MIMEApplication(open(f'{attachment_dir}{file}', 'rb').read())
                attachment.add_header('Content-Disposition', 'attachment', filename=file)
                self.msg.attach(attachment)
                print(f'\n      ATTACHMENTS: {attachment.get_filename()}')

    def send_mail(self):
        with smtplib.SMTP('smtp.office365.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(self.from_email_address, self.smtp_password)
            smtp.send_message(self.msg)

    def eml_to_pdf(self):
        fp = None
        temp = []
        parts = self.msg.get_payload()

        for p in parts:

            ct = p.get_content_type()

            if 'image/jpeg' in ct or 'image/png' in ct:
                continue
            if 'text/html' in ct and 'Content-Disposition: attachment;' not in ct:

                fp = os.path.join(self.default_email_dir, f'Export/{self.directory}.export.html')
                if os.path.dirname(fp):
                    os.makedirs(os.path.dirname(fp), exist_ok=True)
                temp.append(p)

        with open(fp, 'wb') as html:
            html.write(temp[0].get_payload(decode=True))

        pic_list = self.sort_pics()

        with open(fp, encoding='utf-8') as html:
            pic_replacement = html.read()

        for pic in pic_list:
            src = f'cid:image{pic_list.index(pic)}'
            if self.directory == '':
                swap = os.path.join('..\\..\\' + self.default_images_dir,
                                    os.listdir(self.default_images_dir)[pic_list.index(pic)]
                                    )
                pic_replacement = pic_replacement.replace(src, swap)

            else:
                swap = os.path.join('..\\..\\..\\' + self.default_images_dir,
                                    os.listdir(self.default_images_dir)[pic_list.index(pic)]
                                    )
                pic_replacement = pic_replacement.replace(src, swap)
        print('\n   .EML to .HTML: Writing e-mail body to HTML')
        with open(fp, 'w', encoding='utf-8') as html:
            html.write(pic_replacement)

        path_to_wkhtmltopdf = r'"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"' \
                              r' --enable-local-file-access -q'

        path_to_file = fp
        output_path = os.path.join(self.default_email_dir, f'Export/{self.directory}.export.pdf')
        print('\n   .HTML to .PDF: Converting body to PDF')
        os.system(path_to_wkhtmltopdf + ' ' + path_to_file + ' ' + output_path)


def notify(send_email=True, print_to_dir=False):

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

            eml = Mailer(from_email_address, smtp_password, **csv_row)

            if send_email is True:
                print(f'\n*** Sending email to {csv_row["first_name"].upper().strip()} '
                      f'{csv_row["last_name"].upper().strip()} '
                      f'<{csv_row["to_email"].lower().strip()}> ***'

                      )
                eml.send_mail()

            if print_to_dir is True:
                print(f'\n*** Printing email to {eml.default_email_dir}\\Export ***')
                (eml.eml_to_pdf())

            if print_to_dir is False and send_email is False:
                print('\n*** No output generated and no email sent ***')

            print('\n------------------------------------------------------')


def main():
    notify()


if __name__ == '__main__':
    main()
