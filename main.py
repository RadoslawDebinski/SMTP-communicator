import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import imaplib
import email
import os
import sys

from PyQt5.QtWidgets import (
    QApplication, QDialog, QMainWindow, QMessageBox
)
from PyQt5.uic import loadUi

from mainwindow.ui import Ui_MainWindow

filePath = 'C:\\ACIR-WETI\\sem 6\\Współczesne narzędzia obliczeniowe II\\Laby\\Mail\\attachments'

# Dane do logowania
SMTP_SERVER = 'smtp.wp.pl'
SMTP_PORT = 587
SMTP_USERNAME = 'labyWNO@wp.pl'
SMTP_PASSWORD = 'Radoslaww100WNO'

IMAP_SERVER = 'imap.wp.pl'
IMAP_PORT = 993
IMAP_USERNAME = 'labyWNO@wp.pl'
IMAP_PASSWORD = 'Radoslaww100WNO'

# Odczyt maili

def unread_list():
    # Połączenie z serwerem IMAP
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(IMAP_USERNAME, IMAP_PASSWORD)
    mail.select('inbox')

    # Wyszukanie nieprzeczytanych maili
    _, search_data = mail.search(None, 'UNSEEN')
    mail_ids = search_data.split()
    mail_ids = '\n'.join(mail_ids)

    return mail_ids


# def read_email(choice):
#     # Tworzenie nowego katalogu na pliki z załącznikami
#     if not os.path.exists(filePath):
#         os.mkdir(filePath)
#
#     # Połączenie z serwerem IMAP
#     mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
#     mail.login(IMAP_USERNAME, IMAP_PASSWORD)
#     mail.select('inbox')
#
#     # Wyszukanie nieprzeczytanych maili
#     _, search_data = mail.search(None, 'UNSEEN')
#     mail_ids = search_data[choice].split()
#
#     mailData = []
#
#     for i in mail_ids:
#             _, data = mail.fetch(i, '(RFC822)')
#             _, b = data[0]
#             email_message = email.message_from_bytes(b)
#             print('Od: ' + email_message['From'])
#             mailData.append(email_message['From'])
#             print('Temat: ' + email_message['Subject'])
#             mailData.append(email_message['Subject'])
#             if email_message.is_multipart():
#                 for part in email_message.walk():
#                     content_type = part.get_content_type()
#                     if 'image' in content_type:
#                         # Zapisanie załącznika jako pliku
#                         filename = part.get_filename()
#                         if filename:
#                             dir_list = os.listdir(filePath)
#                             lastImgId = int(dir_list[-1][3::-5])
#                             imgPath = os.path.join(filePath, 'img'+str(lastImgId+1)+'.jpg')
#                             print("New attachment added: " + 'img'+str(lastImgId+1)+'.jpg')
#                             with open(imgPath, 'wb') as f:
#                                 f.write(part.get_payload(decode=True))
#                     elif 'text/plain' in content_type:
#                         # Wyświetlenie treści maila jako zwykłego tekstu
#                         print(part.get_payload())
#                         mailData.append(part.get_payload())
#             else:
#                 # Wyświetlenie treści maila jako zwykłego tekstu
#                 print(email_message.get_payload())
#                 mailData.append(email_message.get_payload())
#
#         mail.close()
#         mail.logout()
#
#     return mailData

# Wysyłanie maili
def send_email(to, subject, message, image_path=None):
    # Utworzenie obiektu wiadomości
    msg = MIMEMultipart()
    msg['From'] = SMTP_USERNAME
    msg['To'] = to
    msg['Subject'] = subject

    # Dodanie treści wiadomości
    msg.attach(MIMEText(message, 'plain'))

    # Dodanie załącznika w postaci obrazu
    if image_path:
        with open(image_path, 'rb') as f:
            img_data = f.read()
        image = MIMEImage(img_data, name='image.jpg')
        msg.attach(image)

    # Wysłanie wiadomości
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(SMTP_USERNAME, SMTP_PASSWORD)
    server.sendmail(SMTP_USERNAME, to, msg.as_string())
    server.quit()


class Window(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.connectSignalsSlots()

    def connectSignalsSlots(self):
        self.action_Exit.triggered.connect(self.close)
        self.action_Find_Replace.triggered.connect(self.findAndReplace)
        self.action_About.triggered.connect(self.about)

    def findAndReplace(self):
        dialog = FindReplaceDialog(self)
        dialog.exec()

    def about(self):
        QMessageBox.about(
            self,
            "About Sample Editor",
            "<p>A sample text editor app built with:</p>"
            "<p>- PyQt</p>"
            "<p>- Qt Designer</p>"
            "<p>- Python</p>",
        )


class FindReplaceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        loadUi("ui/find_replace.ui", self)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = Window()
    win.show()
    sys.exit(app.exec())
