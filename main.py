import sys
import imaplib
import smtplib
import email.utils
from email.mime.text import MIMEText

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QTextEdit, \
    QPushButton
import os
from PIL import Image
from datetime import datetime

filePath = '...'
newWidth = 64
resMess = "Pls leave me alone. \n\n Yours sincerely \n\n labywno@wp.pl"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Klient poczty")
        self.setGeometry(100, 100, 800, 600)
        self.centralWidget = QWidget(self)
        self.setCentralWidget(self.centralWidget)
        self.mainLayout = QVBoxLayout()
        self.centralWidget.setLayout(self.mainLayout)
        self.emailLayout = QHBoxLayout()
        self.mainLayout.addLayout(self.emailLayout)
        self.emailLabel = QLabel("Adres email:")
        self.emailLayout.addWidget(self.emailLabel)
        self.emailEdit = QLineEdit()
        self.emailLayout.addWidget(self.emailEdit)
        self.passwordLayout = QHBoxLayout()
        self.mainLayout.addLayout(self.passwordLayout)
        self.passwordLabel = QLabel("Hasło:")
        self.passwordLayout.addWidget(self.passwordLabel)
        self.passwordEdit = QLineEdit()
        self.passwordEdit.setEchoMode(QLineEdit.Password)
        self.passwordLayout.addWidget(self.passwordEdit)
        self.loginButton = QPushButton("Zaloguj")
        self.mainLayout.addWidget(self.loginButton)
        self.loginButton.clicked.connect(self.login)
        self.messagesEdit = QTextEdit()
        self.mainLayout.addWidget(self.messagesEdit)
        self.sendLayout = QHBoxLayout()
        self.mainLayout.addLayout(self.sendLayout)
        self.toLabel = QLabel("Do:")
        self.sendLayout.addWidget(self.toLabel)
        self.toEdit = QLineEdit()
        self.sendLayout.addWidget(self.toEdit)
        self.subjectLabel = QLabel("Temat:")
        self.sendLayout.addWidget(self.subjectLabel)
        self.subjectEdit = QLineEdit()
        self.sendLayout.addWidget(self.subjectEdit)
        self.sendButton = QPushButton("Wyślij")
        self.sendLayout.addWidget(self.sendButton)
        self.sendButton.clicked.connect(self.send)

        self.resLayout = QHBoxLayout()
        self.mainLayout.addLayout(self.resLayout)
        self.fromLabel = QLabel("Responder od:")
        self.resLayout.addWidget(self.fromLabel)
        self.fromEdit = QLineEdit()
        self.resLayout.addWidget(self.fromEdit)
        self.untilLabel = QLabel("Responder do:")
        self.resLayout.addWidget(self.untilLabel)
        self.untilEdit = QLineEdit()
        self.resLayout.addWidget(self.untilEdit)
        self.sendButton = QPushButton("Ustaw")
        self.resLayout.addWidget(self.sendButton)
        self.sendButton.clicked.connect(self.resSet)

    def convert_ACII_art(self, imagePath):
        img = Image.open(imagePath)

        # resize image
        width, height = img.size
        aspectRatio = height / width
        newHeight = aspectRatio * newWidth * 0.55
        img = img.resize((newWidth, int(newHeight)))

        # convert image to greyscale format
        img = img.convert('L')
        pixels = img.getdata()

        # replace each pixel as character from array of printable chars
        chars = ["!", "@", "#", "$", "%", "^", "&", "*", "{", "}", "|"]  # string.printable = all printable chars
        newPixels = [chars[pixel // 25] for pixel in pixels]
        newPixels = ''.join(newPixels)

        # split string of chars into multiple strings of length equal to new width and create a list
        newPixelsCount = len(newPixels)
        asciiImage = [newPixels[index:index + newWidth] for index in range(0, newPixelsCount, newWidth)]
        # send image
        for i in asciiImage:
            self.messagesEdit.append(str(i))

    def resSet(self):

        timer = QTimer(self)
        timer.timeout.connect(self.responder)
        timer.start(10000)

    def responder(self):

        today = datetime.today().strftime('%Y-%m-%d')
        resFrom = self.fromEdit.text()
        resTo = self.untilEdit.text()

        if resFrom < today < resTo:

            _, search_data = self.imap.search(None, 'UNSEEN')
            mail_ids = search_data[0].split()

            if len(mail_ids) != 0:
                for mail_id in mail_ids:
                    _, data = self.imap.fetch(mail_id, '(RFC822)')
                    _, b = data[0]
                    email_message = email.message_from_bytes(b)

                    smtp = smtplib.SMTP('smtp.wp.pl', 587)
                    smtp.ehlo()
                    smtp.starttls()
                    emailToRespond = self.emailEdit.text()
                    password = self.passwordEdit.text()
                    smtp.login(emailToRespond, password)
                    msgR = MIMEText(resMess)
                    msgR['To'] = email_message['From']
                    msgR['Subject'] = "Auto respond"
                    msgR['From'] = email.utils.formataddr(('Sender', emailToRespond))
                    smtp.sendmail(emailToRespond, email_message['From'], msgR.as_string())
                    self.imap.store(mail_id, '+FLAGS', '\Seen')
                    smtp.quit()

    def login(self):
        self.messagesEdit.clear()
        email = self.emailEdit.text()
        password = self.passwordEdit.text()
        self.imap = imaplib.IMAP4_SSL('imap.wp.pl', 993)
        log = False
        try:
            self.imap.login(email, password)  # email, password
            log = True
        except:
            self.messagesEdit.clear()
            self.messagesEdit.append('Invalid login or password')
        if log:
            self.imap.select('INBOX')
            self.fetch_messages()

    def fetch_messages(self):
        _, search_data = self.imap.search(None, 'ALL')
        mail_ids = search_data[0].split()

        for mail_id in mail_ids:
            _, data = self.imap.fetch(mail_id, '(RFC822)')
            _, b = data[0]
            email_message = email.message_from_bytes(b)
            self.messagesEdit.append("\n" + 'Od: ' + str(email_message['From']) + "\n")
            self.messagesEdit.append('Temat: ' + str(email_message['Subject']) + "\n")
            if email_message.is_multipart():
                for part in email_message.walk():
                    content_type = part.get_content_type()
                    if 'image' in content_type:
                        filename = part.get_filename()
                        if filename:
                            dir_list = os.listdir(filePath)
                            lastImgId = int(dir_list[-1][3::-5])
                            imgPath = os.path.join(filePath, 'img' + str(lastImgId + 1) + '.jpg')
                            self.messagesEdit.append(
                                "\n[New attachment added: " + 'img' + str(lastImgId + 1) + '.jpg]\n')
                            with open(imgPath, 'wb') as f:
                                f.write(part.get_payload(decode=True))
                            self.convert_ACII_art(imgPath)

                    elif 'text/plain' in content_type:
                        self.messagesEdit.append(part.get_payload())
            else:
                self.messagesEdit.append(email_message.get_payload())

    def send(self):
        smtp = smtplib.SMTP('smtp.wp.pl', 587)
        smtp.ehlo()
        smtp.starttls()
        emailToSend = self.emailEdit.text()
        password = self.passwordEdit.text()
        smtp.login(emailToSend, password)
        msg = MIMEText(self.messagesEdit.toPlainText())
        msg['To'] = self.toEdit.text()
        msg['Subject'] = self.subjectEdit.text()
        msg['From'] = email.utils.formataddr(('Sender', emailToSend))
        smtp.sendmail(emailToSend, self.toEdit.text(), msg.as_string())
        smtp.quit()
        self.messagesEdit.clear()
        self.toEdit.clear()
        self.subjectEdit.clear()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
