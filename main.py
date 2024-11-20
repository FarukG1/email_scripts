import os
import csv
import configparser
import tkinter as tk
from tkinter import filedialog
from imapclient import IMAPClient


class Colors:
    """Farben für die Ausgabe im Terminal."""

    def __init__(self):
        """Initialisiert die Klasse."""
        self.__dir__ = {
            "PURPLE": "\033[95m",
            "BLUE": "\033[94m",
            "CYAN": "\033[96m",
            "GREEN": "\033[92m",
            "WARNING": "\033[93m",
            "ERROR": "\033[91m",
            "RESET": "\033[0m",
            "BOLD": "\033[1m",
            "UNDERLINE": "\033[4m",
        }

    def print(self, color, text):
        """Fügt Farbe zum Text hinzu."""
        return print(f"{self.__dir__[color]}{text}{self.__dir__["RESET"]}")


class Emails:
    """Emails von einem IMAP-Server abrufen."""

    filename: str = None
    host: str = None
    port: int = None
    ssl: bool = False
    users: list = []

    def __init__(self):
        """Initialisiert die Klasse."""
        config = configparser.ConfigParser()
        config.read("./script.ini")
        self.filename = config["SETTINGS"]["Filename"]
        self.host = config["SETTINGS"]["Host"]
        self.port = int(config["SETTINGS"]["Port"])
        self.ssl = bool(config["SETTINGS"]["SSL"])
        for index in range(1, int(config["USERS"]["Count"]) + 1):
            self.users.append(
                [
                    config["USERS"][f"Username{index}"],
                    config["USERS"][f"Password{index}"],
                ]
            )

    def fetch_one_user(self, username, password):
        """Fetch emails from a specific folder of a user."""
        _sent_mails = dict()
        _inbox_mails = dict()
        server = IMAPClient(
            host="imap.ionos.de",
            port=993,
            use_uid=True,
            ssl=True,
        )
        server.login(username, password)
        server.select_folder("Gesendete Objekte")
        _sent_mails = server.fetch(server.search(), ["ENVELOPE"])
        server.close_folder()
        server.select_folder("INBOX")
        _inbox_mails = server.fetch(server.search(), ["ENVELOPE"])
        server.close_folder()
        server.logout()
        return _sent_mails, _inbox_mails


def open_file_explorer():
    """Öffnet den Datei-Explorer und gibt den Pfad zur ausgewählten Datei zurück."""
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename()
    return file_path


# Initialize colors class
colors = Colors()

# Generate initial configuration file
if not os.path.exists("./script.ini"):
    config = configparser.ConfigParser()
    config["SETTINGS"] = {
        "Filename": "",
        "Host": "",
        "Port": "",
        "SSL": "",
    }
    config["USERS"] = {"Count": "1", "Username1": "", "Password1": ""}
    with open("./script.ini", "w", encoding="utf-8") as configfile:
        config.write(configfile)
    colors.print(
        "ERROR",
        "Please fill out the configuration file (script.ini) and run the script again.",
    )
    exit(0)

# Initialize emails class
emails = Emails()

# Fetch emails for each user
sent_addresses = []
inbox_addresses = []
for user in emails.users:
    colors.print("CYAN", f"Fetching emails for {user[0]}...")
    sent_mails, inbox_mails = emails.fetch_one_user(user[0], user[1])

    # Fetch sent emails
    colors.print("CYAN", "Fetching sent emails...")
    for msgid, data in sent_mails.items():
        envelope = data[b"ENVELOPE"]
        address = (
            envelope.to[0].mailbox.decode(encoding="utf-8")
            + "@"
            + envelope.to[0].host.decode(encoding="utf-8")
        )
        sent_addresses.append(address)

    # Fetch inbox emails
    colors.print("CYAN", "Fetching inbox emails...")
    for msgid, data in inbox_mails.items():
        envelope = data[b"ENVELOPE"]
        address = (
            envelope.from_[0].mailbox.decode(encoding="utf-8")
            + "@"
            + envelope.from_[0].host.decode(encoding="utf-8")
        )
        inbox_addresses.append(address)

# Print count of sent and inbox emails
colors.print("GREEN", "Sent Count: " + str(len(set(sent_addresses))))
colors.print("ERROR", "Inbox Count: " + str(len(set(inbox_addresses))))

count = 0
csv_rows = []
for address in set(sent_addresses).difference(set(inbox_addresses)):
    csv_rows.append([address])
    count += 1
with open(emails.filename, "w", newline="", encoding="utf-8") as csvfile:
    writer = csv.writer(
        csvfile, delimiter=",", quotechar="|", quoting=csv.QUOTE_MINIMAL
    )
    writer.writerow(["Email"])
    writer.writerows(csv_rows)
colors.print("CYAN", f"Total No-Replys: {count + 1}")
