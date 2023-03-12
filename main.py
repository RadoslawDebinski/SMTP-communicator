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

from nltk.corpus import stopwords
stop_words = stopwords.words('english')

# words matching model import
import spacy
nlp = spacy.load('en_core_web_md')

# file path to attachements file
filePath = '...'
# width od images in attachments to print
newWidth = 64
# message for autoresponder
resMess = "..."


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # setting login data / buttons
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
        # setting sending data / buttons
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
        # setting responder data / buttons
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
        # setting key word checking data / buttons
        self.findLayout = QHBoxLayout()
        self.mainLayout.addLayout(self.findLayout)
        self.findLabel = QLabel("Słowo kluczowe:")
        self.findLayout.addWidget(self.findLabel)
        self.findEdit = QLineEdit()
        self.findLayout.addWidget(self.findEdit)
        self.sendButton = QPushButton("Szukaj")
        self.findLayout.addWidget(self.sendButton)
        self.sendButton.clicked.connect(self.findWords)

    # conversion image to ASCIIart algorithm
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
        # connect parts of image
        for i in asciiImage:
            self.messagesEdit.append(str(i))

    # preprocessing data for key word checking algorithm = no spaces, split, lower cases
    def preprocess(self, sentence):
        return [w for w in sentence.lower().split() if w not in stop_words]

    # checking similarity between key word and all words in messages algorithm
    def checkSim(self, part, email_message):
        sampleWord = self.findEdit.text()
        # taking data depends on form of message = with attachment or not
        if part is None:
            part = email_message
            mess = email_message.get_payload()
        else:
            mess = part.get_payload()
        # preprocess message
        messP = self.preprocess(mess)
        simData = []
        for i in messP:
            tokens = nlp(f"{sampleWord} {i}")
            token1, token2 = tokens[0], tokens[1]
            sim = round(token1.similarity(token2) * 100)
            # acceptable level of similarity in %
            if sim > 60:
                simData.append(f"{i} = {sim}%")
        if len(simData) > 0:
            self.messagesEdit.append("\n" + 'Od: ' + str(email_message['From']) + "\n")
            self.messagesEdit.append('Temat: ' + str(email_message['Subject']) + "\n")
            self.messagesEdit.append(part.get_payload())
            self.messagesEdit.append(f"\n{simData}\n")

    # connection to email to preprocess messages
    def findWords(self):
        self.messagesEdit.clear()
        _, search_data = self.imap.search(None, 'ALL')
        mail_ids = search_data[0].split()

        for mail_id in mail_ids:
            _, data = self.imap.fetch(mail_id, '(RFC822)')
            _, b = data[0]
            email_message = email.message_from_bytes(b)

            if email_message.is_multipart():
                for part in email_message.walk():
                    content_type = part.get_content_type()
                    if 'image' in content_type:
                        pass
                    elif 'text/plain' in content_type:
                        self.checkSim(part, email_message)
            else:
                self.checkSim(None, email_message)

    # setting timer for responder
    def resSet(self):

        timer = QTimer(self)
        timer.timeout.connect(self.responder)
        timer.start(10000)

    def responder(self):
        # gathering time intervals
        today = datetime.today().strftime('%Y-%m-%d')
        resFrom = self.fromEdit.text()
        resTo = self.untilEdit.text()
        # checking time intervals
        if resFrom < today < resTo:

            _, search_data = self.imap.search(None, 'UNSEEN')
            mail_ids = search_data[0].split()
            # auto respond for all unseen messages and changing their status
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

    # login to email
    def login(self):
        self.messagesEdit.clear()
        # gathering data from user
        email = self.emailEdit.text()
        password = self.passwordEdit.text()
        self.imap = imaplib.IMAP4_SSL('imap.wp.pl', 993)
        log = False
        # checking data are acceptable
        try:
            self.imap.login(email, password)  # email, password
            log = True
        # there should be external error of email
        except:
            self.messagesEdit.clear()
            self.messagesEdit.append('Invalid login or password')
        if log:
            self.imap.select('INBOX')
            self.fetch_messages()

    # messages analysis and conversion of images in attachments
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

    # sending email by data gathered form user
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


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
