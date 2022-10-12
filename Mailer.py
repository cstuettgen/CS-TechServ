import pathlib
import smtplib
import sys
import os
import csv
import winreg
import logging
import shutil
import importlib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from weasyprint import HTML
from datetime import datetime
from PySide6.QtCore import (QMetaObject, QRect)
from PySide6.QtWidgets import (QApplication, QGroupBox, QLabel, QLineEdit, QPushButton, QMainWindow)


class Mailer:
    def __init__(self, **kwargs):
        self.from_email_address = get_registry('Mailer.smtp')
        self.smtp_password = get_registry('Mailer.pass')
        self.directory = kwargs['directory']
        self.first_name = kwargs['first_name']
        self.last_name = kwargs['last_name']

        self.carbon_copy = kwargs['carbon_copy']
        self.subject = kwargs['subject']
        self.msg = None
        self.date = datetime.now().strftime("%Y-%m-%d %H_%M_%S")
        if self.directory != 'log':
            self.to_email = kwargs['to_email']
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
        with open('lastrun.log', "r") as last_run:
            log_data = last_run.read()

        if self.directory == 'log':
            msg = MIMEMultipart('related')
            msg['From'] = self.from_email_address
            msg['To'] = self.from_email_address
            msg['cc'] = self.carbon_copy
            msg['Subject'] = self.subject
            msg.attach(MIMEText(log_data))
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
            logging.info('')
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

            logging.info(f"TO: {self.to_email} ")
            logging.info(f"FROM: {self.from_email_address} ")
            logging.info(f"CC: {self.carbon_copy} ")

            for item in kwargs:
                f_item = '{'+item+'}'

                if f_item in self.subject:
                    self.subject = self.subject.replace(f_item, kwargs[item])
                if f_item in formatted:
                    formatted = formatted.replace(f_item, kwargs[item])

            msg['Subject'] = self.subject
            logging.info(f'SUBJECT: {self.subject}\n')
            if self.directory == '':
                logging.info(f'*** Reading HTML file, "{self.default_email_dir}email.html", into email body... ***')
            else:
                logging.info(f'*** Reading HTML file,'
                             f' "{self.default_email_dir}\\{self.directory}.html", into email body... ***')
            logging.debug(f'----------------------- Email body .HTML: ---------------------------'
                          f'\n{formatted}\n'
                          f'-------------------------------------------------------------------------------'
                          f'------------')
            logging.info('      Attaching Message')
            msg.attach(MIMEText(formatted, 'html'))

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
            logging.info(f'*** Sending log file to {self.from_email_address} ***')

        with smtplib.SMTP('smtp.office365.com', 587) as smtp:
            importlib.reload(os)
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

        logging.info(f'*** Converting Email Body to HTML/PDF ***')
        logging.info(f'   .EML to .HTML: Writing e-mail body to {fp}')

        with open(fp, 'w', encoding='utf-8') as html:
            html.write(pic_replacement)

        if self.directory == '':
            output_path = os.path.join(self.default_email_dir, f'Export\\email.export.pdf')
        else:
            output_path = os.path.join(self.default_email_dir, f'Export\\{self.directory}.export.pdf')

        logging.info(f'   .HTML to .PDF: Converting body to {output_path}')

        HTML(fp).write_pdf(output_path)

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


def notify(send_email=False, write_to_dir=False, attach_email_as_pdf=False, log_level=20,
           send_log=False, change_creds=False):

    logger(log_level)

    if change_creds is True or get_registry('Mailer.smtp') == '' or get_registry('Mailer.pass') == '':
        creds()

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

            eml = Mailer(**csv_row)

            if write_to_dir is True:
                save_pdf(eml)
            if attach_email_as_pdf is True:
                attach(eml, send_email, write_to_dir)
            if send_email is True:
                send(eml, csv_row)
            if write_to_dir is False and send_email is False:
                logging.info('*** No output generated and no email sent ***')

    logging.info('\n\n --- - --- - --- - --- - ======= END OF RUN ======= - --- - --- - --- - ---\n')
    if send_log is True:
        log_mail()

    append_log()


def attach(eml, send_email, write_to_dir):
    if send_email is False:
        logging.info(' !======!     SEND EMAIL NOT SELECTED!    !======!')
    else:
        eml.attach_pdf(write_to_dir)


def send(eml, csv_row):
    logging.info(f'*** Sending email to {csv_row["first_name"].upper().strip()} '
                 f'{csv_row["last_name"].upper().strip()} '
                 f'<{csv_row["to_email"].lower().strip()}> ***'
                 )
    eml.send_mail()


def save_pdf(eml):
    pdf_path = eml.eml_to_pdf()
    logging.info(f'*** Email saved to {pdf_path} ***')


def creds():
    app = QApplication(sys.argv)
    window = UiForm()
    window.show()
    app.exec()
    app.quit()


def log_mail():
    kwargs = {'first_name': 'MAILER', 'last_name': 'LOG', 'carbon_copy': '',
              'subject': 'MAILER: Log from last run', 'directory': 'log'}
    mail = Mailer(**kwargs)
    mail.send_mail()


def logger(log_level):
    logging.basicConfig(format='%(asctime)s  %(levelname)-8s %(message)s',
                        level=log_level,
                        datefmt='%Y-%m-%d %H:%M:%S',
                        filename=r'lastrun.log',
                        filemode='w',
                        )
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    formatter = logging.Formatter('%(levelname)-8s %(message)s')
    console.setFormatter(formatter)
    logging.getLogger().addHandler(console)

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


def set_registry(keyname, keyvalue, regdir='Environment',):
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, regdir) as _:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, regdir, 0, winreg.KEY_WRITE) as writeRegistryDir:
            winreg.SetValueEx(writeRegistryDir, keyname, 0, winreg.REG_SZ, keyvalue)
    return f'HKCU/{regdir}/{keyname}/', keyvalue


def get_registry(keyname, regdir='Environment'):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, regdir) as accessRegistryDir:
            value, _ = winreg.QueryValueEx(accessRegistryDir, keyname)
    except FileNotFoundError:
        set_registry(keyname, '')
        value = get_registry(keyname)
    return value


class UiForm(QMainWindow):
    def __init__(self):
        super(UiForm, self).__init__()
        self.resize(500, 400)
        QMetaObject.connectSlotsByName(self)
        self.groupbox = QGroupBox(self)
        self.groupbox.setObjectName(u"groupBox")
        self.groupbox.setGeometry(QRect(10, 20, 321, 181))
        self.username = QLineEdit(self.groupbox)
        self.username.setObjectName(u"username")
        self.username.setGeometry(QRect(10, 40, 301, 22))
        self.password = QLineEdit(self.groupbox)
        self.password.setObjectName(u"password")
        self.password.setEchoMode(QLineEdit.Password)
        self.password.setGeometry(QRect(10, 110, 301, 22))
        self.label = QLabel(self.groupbox)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(10, 20, 161, 16))
        self.label_2 = QLabel(self.groupbox)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(10, 90, 49, 16))
        self.pushbutton = QPushButton(self)
        self.pushbutton.setObjectName(u"pushButton")
        self.pushbutton.setGeometry(QRect(140, 220, 75, 24))
        self.setWindowTitle("Set SMTP Credentials")
        self.groupbox.setTitle("SMTP Credentials:")
        self.label.setText("Username:")
        self.label_2.setText("Password:")
        self.pushbutton.setText("OK")
        self.pushbutton.clicked.connect(self.click)

    def click(self):
        logging.info('*** Storing credentials ***')
        username = set_registry('Mailer.smtp', self.username.text())
        logging.info(f'SMTP Username, {username[0]} [{username[1]}], set')
        password = set_registry('Mailer.pass', self.password.text())
        logging.info(f'SMTP Password, {password[0]} [***********], set\n')
        self.close()


if __name__ == '__main__':
    main()
