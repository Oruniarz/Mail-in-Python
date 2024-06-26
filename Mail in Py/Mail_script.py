import smtplib
import email
import imaplib
from email.header import decode_header, Header
import os
import sys
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QTextEdit, QPushButton
from PySide6.QtGui import QPalette
from functools import partial
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from sentence_transformers import CrossEncoder
model = CrossEncoder("cross-encoder/ms-marco-MiniLM-L-6-v2")


MY_EMAIL = os.environ["MY_EMAIL"]
PASSWORD = os.environ["PASSWORD"]
SERVER = 'imap.gmail.com'
AUTORESPONDER_MESSAGE = """
Jestem teraz na wyjeździe. Odczytam wiadomość i w miarę możliwości odpowiem po 31 marca.

Z poważaniem,
Ja
"""


class Poczta(QWidget):
    def __init__(self):
        super().__init__()
        self.w = None
        self.respond = False

        self.setWindowTitle('Poczta')
        okno_glowne = QVBoxLayout()
        self.setLayout(okno_glowne)

        czesc_email = QHBoxLayout()
        email_label = QLabel("Adresat:")
        self.email_content = QLineEdit()
        czesc_email.addWidget(email_label)
        czesc_email.addWidget(self.email_content)
        okno_glowne.addLayout(czesc_email)

        czesc_tematu = QHBoxLayout()
        temat_label = QLabel("Temat:")
        self.temat_content = QLineEdit()
        czesc_tematu.addWidget(temat_label)
        czesc_tematu.addWidget(self.temat_content)
        okno_glowne.addLayout(czesc_tematu)

        czesc_wiadomosci = QLabel("Wiadomość:")
        self.wiadomosc_content = QTextEdit()
        okno_glowne.addWidget(czesc_wiadomosci)
        okno_glowne.addWidget(self.wiadomosc_content)

        wyslij_wiadmomosc_button = QPushButton("Wyślij")
        wyslij_wiadmomosc_button.clicked.connect(self.wyslij_wiadomosc)
        okno_glowne.addWidget(wyslij_wiadmomosc_button)

        odczyt_wiadomosci_button = QPushButton("Odczyt")
        odczyt_wiadomosci_button.clicked.connect(self.odczyt_wiadomosci)
        okno_glowne.addWidget(odczyt_wiadomosci_button)

        self.autoresponder_button = QPushButton("Autoresponder wyłączony")
        self.autoresponder_button.clicked.connect(self.autoresponder)
        okno_glowne.addWidget(self.autoresponder_button)

    def wyslij_wiadomosc(self):
        if self.email_content.text():
            msg = MIMEMultipart()
            msg["From"] = MY_EMAIL
            msg["To"] = self.email_content.text()
            msg["Subject"] = Header(self.temat_content.text(), "utf-8").encode()
            msg.attach(MIMEText(self.wiadomosc_content.toPlainText(), "plain", "utf-8"))
            with smtplib.SMTP("smtp.gmail.com", 587) as connection:
                connection.starttls()
                connection.login(user=MY_EMAIL, password=PASSWORD)
                connection.send_message(msg)
                print("Wiadomość wysłana!")

    def odczyt_wiadomosci(self):
        self.w = Wiadomosci(self.zaladuj_wiadomosci(), self.respond)
        self.w.show()

    def autoresponder(self):
        if not self.respond:
            self.respond = True
            self.autoresponder_button.setText("Autoresponder włączony")
        else:
            self.respond = False
            self.autoresponder_button.setText("Autoresponder wyłączony")

    def zaladuj_wiadomosci(self):
        email_list = []
        try:
            mail = imaplib.IMAP4_SSL(SERVER)
            mail.login(MY_EMAIL, PASSWORD)
            mail.select("inbox")
            status, messages = mail.search(None, "ALL")

            for response in messages[0].split():
                status, messages = mail.fetch(response, '(RFC822)')
                msg = email.message_from_bytes(messages[0][1])
                content = {}

                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding)
                content["subject"] = subject
                from_address, encoding = decode_header(msg.get("From"))[0]
                if isinstance(from_address, bytes):
                    from_address = from_address.decode(encoding)
                content["from"] = from_address
                content["body"] = self.get_email_body(msg)
                email_list.append(content)

            mail.logout()

        except Exception as e:
            print(f"Error: {str(e)}")

        return email_list

    def get_email_body(self, msg):
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == 'text/plain':
                    body += part.get_payload(decode=True).decode()
        else:
            body = msg.get_payload(decode=True).decode()
        return body


class Wiadomosci(QWidget):
    def __init__(self, email_list, respond):
        super().__init__()
        self.email_list = email_list
        self.win = None
        self.respond = respond
        self.button_palette = None

        self.setWindowTitle("Wiadomości")

        layout = QVBoxLayout()
        self.setLayout(layout)

        query = QHBoxLayout()
        self.query_box = QLineEdit()
        search_button = QPushButton("Szukaj")
        search_button.clicked.connect(self.wyszukaj)
        query.addWidget(self.query_box)
        query.addWidget(search_button)
        layout.addLayout(query)

        self.buttons = []
        for mail in self.email_list:
            button = QPushButton(f"Od: {mail['from']}\nTemat: {mail['subject']}")
            button.clicked.connect(partial(self.wyswietl_wiadomosc, email=mail))
            layout.addWidget(button)
            self.buttons.append(button)
        self.button_palette = self.buttons[0].palette()

    def wyswietl_wiadomosc(self, email):
        if self.respond:
            msg = MIMEMultipart()
            msg["From"] = MY_EMAIL
            msg["To"] = email["from"]
            msg["Subject"] = Header("Autoresponder", "utf-8").encode()
            msg.attach(MIMEText(AUTORESPONDER_MESSAGE, "plain", "utf-8"))
            with smtplib.SMTP("smtp.gmail.com", 587) as connection:
                connection.starttls()
                connection.login(user=MY_EMAIL, password=PASSWORD)
                connection.send_message(msg)
                print("Wiadomość wysłana!")
        self.win = OknoWiadomosci(email)
        self.win.show()
        pass

    def wyszukaj(self):
        query = self.query_box.text()
        for i, mail in enumerate(self.email_list):
            if model.rank(query, list(mail.values()))[0]["score"] > -0.5:
                self.buttons[i].setStyleSheet("background-color: lightgray; color: blue;")
            else:
                self.buttons[i].setStyleSheet(f"background-color: "
                                              f"{self.button_palette.color(QPalette.Button)}; color: black;")


class OknoWiadomosci(QWidget):
    def __init__(self, mail):
        super().__init__()

        self.setWindowTitle('Wiadomość')
        self.setGeometry(100, 100, 300, 50)
        okno_glowne = QVBoxLayout()
        self.setLayout(okno_glowne)

        from_address = QLabel(f"Od: {mail['from']}")
        subject = QLabel(f"Temat: {mail['subject']}")
        content = QLabel(f"{mail['body']}")

        okno_glowne.addWidget(from_address)
        okno_glowne.addWidget(subject)
        okno_glowne.addWidget(content)


apka = QApplication(sys.argv)
window = Poczta()
window.show()
sys.exit(apka.exec())
