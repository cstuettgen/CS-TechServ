import pathlib
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import shutil
import os
import csv
from weasyprint import HTML
import logging
from _datetime import datetime


class Mailer:
    def __init__(self, from_email_address, smtp_password, **kwargs):
        self.from_email_address = from_email_address
        self.smtp_password = smtp_password
        self.directory = kwargs['directory']
        self.first_name = kwargs['first_name']
        self.last_name = kwargs['last_name']
        self.to_email = kwargs['to_email']
        self.carbon_copy = kwargs['carbon_copy']
        self.subject = kwargs['subject']
        self.msg = None
        self.date = datetime.now().strftime("%Y-%m-%d %H_%M_%S")
        if self.directory != 'log':
            self.default_email_dir = 'Email_Message\\'
            self.default_attachments_dir = 'Attachments\\'
            self.default_images_dir = 'Images\\'
            self.set_dirs()
            self.compose_mail(**kwargs)
            self.embed_pics()
            self.attachments()
        else:
            self.log()

    def log(self):
        lyrics = """What rolls down stairs

Alone or in pairs,

Rolls over your neighbor's dog?

What's great for a snack

and fits on your back?

It's Log, Log, Log!

It's Lo-og, Lo-og,

It's big, it's heavy, it's wood.

It's Lo-og, Lo-og,

It's better than bad, it's good!

Everyone wants a Log!

You're gonna love it, Log!

Come on and get your Log!

Everyone needs a Log!

You're gonna love it, Log! """

        if self.directory == 'log':
            msg = MIMEMultipart('related')
            msg['From'] = self.from_email_address
            msg['To'] = self.to_email
            msg['cc'] = self.carbon_copy
            msg['Subject'] = self.subject
            msg.attach(MIMEText(lyrics))
            attachment = MIMEApplication(open('lastrun.log', 'rb').read())
            attachment.add_header('Content-Disposition', 'attachment', filename=f'Mailer-Last_Run-{self.date}.log')
            msg.attach(attachment)
            self.msg = msg

    def set_dirs(self):
        if self.directory == '':
            logging.info('\n')
            logging.info(f'___________________________Default Directory_________________________________')
            logging.info(f'*** Using default folders for email body, embedded images and attachments ***')

        elif (self.directory not in os.listdir(self.default_email_dir)
                or self.directory not in os.listdir(self.default_images_dir)
                or self.directory not in os.listdir(self.default_attachments_dir)):
            logging.info('\n')
            logging.info(f'___________________________{self.directory}________________________________')
            logging.info(f'*** One or more folders is missing for recipient {self.first_name} {self.last_name}'
                         f' <{self.to_email}> ***')
            exit()

        elif (self.directory in os.listdir(self.default_email_dir)
                and self.directory in os.listdir(self.default_images_dir)
                and self.directory in os.listdir(self.default_attachments_dir)):
            self.default_attachments_dir = os.path.join(self.default_attachments_dir, self.directory)
            self.default_images_dir = os.path.join(self.default_images_dir, self.directory)
            self.default_email_dir = os.path.join(self.default_email_dir, self.directory)
            logging.info('\n')
            logging.info(f'___________________________{self.directory}________________________________')
            logging.info(f'*** Using folder, "{self.directory}"'
                         f', for email body, embedded images and attachments ***')

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
            if self.directory == '':
                logging.info(f'*** Reading HTML file, "{self.default_email_dir}\\email.html", into email body... ***')
            else:
                logging.info(f'*** Reading HTML file,'
                             f' "{self.default_email_dir}\\{self.directory}.html", into email body... ***')
            logging.info(f"TO: {self.to_email} ")
            logging.info(f"FROM: {self.from_email_address} ")
            logging.info(f"CC: {self.carbon_copy} ")

            for item in kwargs:
                f_item = '{'+item+'}'

                if f_item in self.subject:
                    self.subject = self.subject.replace(f_item, kwargs[item])
                if f_item in formatted:
                    formatted = formatted.replace(f_item, kwargs[item])

            logging.debug(f'----------------------- Email body .HTML: ---------------------------'
                          f'\n{formatted}\n'
                          f'-------------------------------------------------------------------------------'
                          f'------------')

            msg['Subject'] = self.subject
            msg.attach(MIMEText(formatted, 'html'))
            logging.info(f'SUBJECT: {self.subject} ')
            self.msg = msg

            return msg

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
        logging.info(f"*** Embedding images from {self.default_images_dir} ***")
        for file in pic_list:
            with open(os.path.join(self.default_images_dir, file), 'rb') as img:
                image = MIMEImage(img.read())

            image.add_header('content-ID', f'<image{i}>')
            self.msg.attach(image)
            logging.info(f'      EMBEDDED IMAGES: {file}')
            i = i + 1

        return pic_list

    def attachments(self):
        attachment_dir = self.default_attachments_dir + '\\'
        attachments = os.listdir(attachment_dir)

        for file in attachments:
            if os.path.isfile(f'{attachment_dir}{file}') is True:
                attachment = MIMEApplication(open(f'{attachment_dir}{file}', 'rb').read())
                attachment.add_header('Content-Disposition', 'attachment', filename=file)
                self.msg.attach(attachment)
                logging.info(f'      ATTACHMENTS: {attachment.get_filename()}')

    def send_mail(self):
        if self.directory == 'log':
            logging.info(f'*** Sending log file to {self.from_email_address}')

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
                if self.directory != '':
                    fp = os.path.join(self.default_email_dir, f'Export\\{self.directory}.export.html')
                else:
                    fp = os.path.join(self.default_email_dir, f'Export\\email.export.html')
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

        logging.info(f'   .EML to .HTML: Writing e-mail body to {fp}')

        with open(fp, 'w', encoding='utf-8') as html:
            html.write(pic_replacement)

        if self.directory == '':
            output_path = os.path.join(self.default_email_dir, f'Export\\email.export.pdf')
        else:
            output_path = os.path.join(self.default_email_dir, f'Export\\{self.directory}.export.pdf')

        logging.info(f'   .HTML to .PDF: Converting body to {output_path}')

        HTML(fp).write_pdf(output_path)

        return output_path

    def attach_pdf(self, write_to_dir):
        if self.directory == 'log':
            pass
        if self.directory == '':
            output_path = os.path.join(self.default_attachments_dir, f'email.export.pdf')
            source_path = os.path.join(self.default_email_dir, f'Export\\email.export.pdf')
            file = f'email.export.pdf'
        else:
            output_path = os.path.join(self.default_attachments_dir, f'{self.directory}.export.pdf')
            source_path = os.path.join(self.default_email_dir, f'Export\\{self.directory}.export.pdf')
            file = f'{self.directory}.export.pdf'
        if write_to_dir is True:
            shutil.copyfile(source_path, output_path)
        else:
            self.eml_to_pdf()
            shutil.copyfile(source_path, output_path)
        attachment = MIMEApplication(open(output_path, 'rb').read())
        attachment.add_header('Content-Disposition', 'attachment', filename=file)
        self.msg.attach(attachment)
        os.remove(output_path)
        logging.info('*** Attaching copy of email body as PDF ***')


def notify(send_email=True, write_to_dir=False, attach_email_as_pdf=False, log_level=20, send_log=False):
    logger(log_level)

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
            logging.debug(f'recipients.csv:\n{header}\n{row}')

            i = 0
            csv_row = {}
            for column in header:
                csv_row.update({column.split()[0]: row[i]})
                i += 1
            eml = Mailer(from_email_address, smtp_password, **csv_row)

            if write_to_dir is True:
                pdf_path = eml.eml_to_pdf()
                logging.info(f'*** Email saved to {pdf_path} ***')
            if attach_email_as_pdf is True:
                if send_email is False:
                    logging.info(' !======!     SEND EMAIL NOT SELECTED!    !======!')
                else:
                    eml.attach_pdf(write_to_dir)
            if send_email is True:
                logging.info(f'*** Sending email to {csv_row["first_name"].upper().strip()} '
                             f'{csv_row["last_name"].upper().strip()} '
                             f'<{csv_row["to_email"].lower().strip()}> ***'
                             )
                eml.send_mail()
            if write_to_dir is False and send_email is False:
                logging.info('*** No output generated and no email sent ***')
    logging.info('                         ======= END OF RUN =======                  \n')
    if send_log is True:
        log_mail(from_email_address, smtp_password)
    with open('lastrun.log') as log:
        for line in log:
            stripped_line = line.strip()
            line_list = stripped_line.split()
            if 'INFO' in line_list:
                print(" ".join(line_list[1:None]))

    append_log()


def log_mail(from_email_address, smtp_password):
    kwargs = {'first_name': 'MAILER', 'last_name': 'LOG', 'to_email': from_email_address,
              'carbon_copy': '', 'subject': 'MAILER: Log from last run', 'directory': 'log'}
    mail = Mailer(from_email_address, smtp_password, **kwargs)
    mail.send_mail()


def logger(log_level):
    logging.basicConfig(format='%(asctime)s  %(levelname)-8s %(message)s',
                        level=log_level,
                        datefmt='%m-%d %H:%M',
                        filename=r'lastrun.log',
                        filemode='w')
    logging.getLogger("weasyprint").setLevel(logging.CRITICAL)
    logging.getLogger("fontTools.subset").setLevel(logging.CRITICAL)
    logging.getLogger("fontTools.ttLib.ttFont").setLevel(logging.CRITICAL)
    logging.getLogger("PIL.PngImagePlugin").setLevel(logging.CRITICAL)


def append_log():
    with open('lastrun.log') as log:
        log_data = log.read()
    with open('Mailer.log', 'a') as log:
        log.write(log_data)
 
 
def main():
    notify()


if __name__ == '__main__':
    main()
