"""Send emails via SMTP."""

import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from smtplib import SMTP_SSL

from imapclient import IMAPClient


class Send:
    """Send emails via SMTP."""

    def __init__(self, config: dict):
        """Initializes the class."""
        self.smtp = {
            "host": config["SMTP"]["host"],
            "port": int(config["SMTP"]["port"]),
            "ssl": bool(config["SMTP"]["ssl"]),
            "sent_folder": config["SMTP"]["sent_folder"],
            "email": config["SMTP"]["email"],
            "username": config["SMTP"]["username"],
            "password": config["SMTP"]["password"],
        }

        self.imap = {
            "host": config["IMAP"]["host"],
            "port": int(config["IMAP"]["port"]),
            "ssl": bool(config["IMAP"]["ssl"]),
            "sent_folder": config["IMAP"]["sent_folder"],
            "inbox_folder": config["IMAP"]["inbox_folder"],
        }

    def create(self, to, subject: str, body: str):
        """Create an email."""
        msg = MIMEMultipart()
        msg["From"] = self.smtp["email"]
        msg["To"] = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))
        return msg

    def send(self, to: str, msg: MIMEMultipart):
        """Send an email."""
        server = SMTP_SSL(self.smtp["host"], self.smtp["port"])
        server.login(self.smtp["username"], self.smtp["password"])
        server.sendmail(self.smtp["email"], to, msg.as_string())
        server.quit()
        return True

    def send_bulk(self, csv_file: str, subject: str, body: str):
        """Send bulk emails."""
        with open(csv_file, "r", encoding="utf-8") as file:
            csv_reader = csv.reader(file)
            email_index = 0

            # Get headers
            for row in csv_reader:
                csv_header = row
                break

            # Get email header index
            for index, column in enumerate(csv_header):
                if "email" in column.lower():
                    email_index = index
                    break

            for row in csv_reader:
                to = row[email_index]
                msg = self.create(to, subject, body)
                self.send(to, msg)

    def save(self, msg: MIMEMultipart):
        """Save an email to the sent folder."""
        server = IMAPClient(self.imap["host"])
        server.login(self.smtp["username"], self.smtp["password"])
        server.select_folder(self.imap["sent_folder"])
        server.append(self.imap["sent_folder"], msg.as_string(), flags=[b"\\Seen"])
        server.close_folder()
        server.logout()
        return True
