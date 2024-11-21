import os
import csv
import configparser
import tkinter as tk
from tkinter import filedialog
from imapclient import IMAPClient


class Colors:
    """Farben für die Ausgabe im Terminal."""

    purple = "\033[95m"
    blue = "\033[94m"
    cyan = "\033[96m"
    green = "\033[92m"
    warning = "\033[93m"
    error = "\033[91m"
    reset = "\033[0m"
    bold = "\033[1m"
    underline = "\033[4m"

    def __init__(self):
        """Initialisiert die Klasse."""
        self.__dir__ = {
            "purple": "\033[95m",
            "blue": "\033[94m",
            "cyan": "\033[96m",
            "green": "\033[92m",
            "warning": "\033[93m",
            "error": "\033[91m",
            "reset": "\033[0m",
            "bold": "\033[1m",
            "underline": "\033[4m",
        }
        self.purple = "\033[95m"
        self.blue = "\033[94m"
        self.cyan = "\033[96m"
        self.green = "\033[92m"
        self.warning = "\033[93m"
        self.error = "\033[91m"
        self.reset = "\033[0m"
        self.bold = "\033[1m"
        self.underline = "\033[4m"

    def print(self, color, text):
        """Fügt Farbe zum Text hinzu."""
        return print(f"{self.__dir__[color]}{text}{self.__dir__["reset"]}")

    def colored(self, color, text):
        """Fügt Farbe zum Text hinzu."""
        return f"{self.__dir__[color]}{text}{self.__dir__["reset"]}"


class Emails:
    """Emails von einem IMAP-Server abrufen."""

    host: str = None
    port: int = None
    ssl: bool = False
    users: list = []

    def __init__(self):
        """Initialisiert die Klasse."""
        config = configparser.ConfigParser()
        config.read("./script.ini")
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


def main():
    """Main function."""
    # Initialize colors class
    colors = Colors()

    # Print ASCII Art
    print(
        rf"""
{colors.purple}$$$$$$$$\                         $$\ $$\        {colors.warning}$$$$$$\                      $$\             $$\
{colors.purple}$$  _____|                        \__|$$ |      {colors.warning}$$  __$$\                     \__|            $$ |
{colors.purple}$$ |      $$$$$$\$$$$\   $$$$$$\  $$\ $$ |      {colors.warning}$$ /  \__| $$$$$$$\  $$$$$$\  $$\  $$$$$$\  $$$$$$\
{colors.purple}$$$$$\    $$  _$$  _$$\  \____$$\ $$ |$$ |      {colors.warning}\$$$$$$\  $$  _____|$$  __$$\ $$ |$$  __$$\ \_$$  _|
{colors.purple}$$  __|   $$ / $$ / $$ | $$$$$$$ |$$ |$$ |       {colors.warning}\____$$\ $$ /      $$ |  \__|$$ |$$ /  $$ |  $$ |
{colors.purple}$$ |      $$ | $$ | $$ |$$  __$$ |$$ |$$ |      {colors.warning}$$\   $$ |$$ |      $$ |      $$ |$$ |  $$ |  $$ |$$\
{colors.purple}$$$$$$$$\ $$ | $$ | $$ |\$$$$$$$ |$$ |$$ |      {colors.warning}\$$$$$$  |\$$$$$$$\ $$ |      $$ |$$$$$$$  |  \$$$$  |
{colors.purple}\________|\__| \__| \__| \_______|\__|\__|       {colors.warning}\______/  \_______|\__|      \__|$$  ____/    \____/
                                                                                  {colors.warning}$$ |
                                                                                  {colors.warning}$$ |
                                                                                  {colors.warning}\__|{colors.reset}""",
    )

    # Generate initial configuration file
    if not os.path.exists("./script.ini"):
        config = configparser.ConfigParser()
        config["SETTINGS"] = {
            "host": "",
            "port": "",
            "ssl": "",
        }
        config["USERS"] = {"count": "1", "username1": "", "password1": ""}
        with open("./script.ini", "w", encoding="utf-8") as configfile:
            config.write(configfile)
        colors.print(
            "error",
            "Please fill out the configuration file (script.ini) and run the script again.",
        )
        return 0

    # Initialize emails class
    emails = Emails()

    # Fetch emails for each user
    sent_addresses = []
    inbox_addresses = []
    for user in emails.users:
        colors.print("cyan", f"Fetching emails for {user[0]}...")
        sent_mails, inbox_mails = emails.fetch_one_user(user[0], user[1])

        # Fetch sent emails
        # colors.print("cyan", "Fetching sent emails...")
        for _, data in sent_mails.items():
            envelope = data[b"ENVELOPE"]
            address = (
                envelope.to[0].mailbox.decode(encoding="utf-8")
                + "@"
                + envelope.to[0].host.decode(encoding="utf-8")
            )
            sent_addresses.append(address)

        # Fetch inbox emails
        # colors.print("cyan", "Fetching inbox emails...")
        for _, data in inbox_mails.items():
            envelope = data[b"ENVELOPE"]
            address = (
                envelope.from_[0].mailbox.decode(encoding="utf-8")
                + "@"
                + envelope.from_[0].host.decode(encoding="utf-8")
            )
            inbox_addresses.append(address)

    # all email adresses to lowercase
    sent_addresses = [address.lower() for address in sent_addresses]
    inbox_addresses = [address.lower() for address in inbox_addresses]

    # Print count of sent and inbox emails
    print("\n")
    colors.print(
        "purple", "Sent Email Addresses Count: " + str(len(set(sent_addresses)))
    )
    colors.print(
        "purple", "Inbox Email Addresses Count: " + str(len(set(inbox_addresses)))
    )

    count = 0
    csv_rows = []
    for address in set(sent_addresses).difference(set(inbox_addresses)):
        csv_rows.append([address])
        count += 1
    colors.print("purple", f"Total No-Replys Email Addresses: {count + 1}\n\n")

    selected_file = open_file_explorer()

    if not selected_file:
        colors.print("error", "No file selected.")
        input(colors.colored("error", "Press Enter to exit and try again..."))
        return 0

    colors.print(
        "green",
        f'Selected File: "\x1b]8;;file://{selected_file}/\x1b\\{selected_file}\x1b]8;;\x1b\\"',
    )

    colors.print("green", "Reading CSV file...")
    csv_rows = []
    with open(selected_file, "r", encoding="utf-8") as csvfile:
        csv_reader = csv.reader(csvfile)
        email_index = 0
        for row in csv_reader:
            for index, column in enumerate(row):
                if "email" in column.lower():
                    email_index = index
                    break
            break

        for row in csv_reader:
            if row[email_index] in set(sent_addresses).difference(set(inbox_addresses)):
                csv_rows.append(row)

    filename = selected_file.split(".csv")[0] + "_no_reply.csv"
    colors.print(
        "green",
        f'Writing CSV file to "\x1b]8;;file://{filename}/\x1b\\{filename}\x1b]8;;\x1b\\"',
    )
    with open(filename, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(
            csvfile, delimiter=",", quotechar="|", quoting=csv.QUOTE_MINIMAL
        )
        writer.writerow(
            [
                "Name",
                "Handle",
                "Link",
                "Followers",
                "Gender",
                "Email",
                "Location",
                "",
                "",
                "",
            ]
        )
        writer.writerows(csv_rows)

    print("\n")
    input(colors.colored("green", "Press Enter to exit..."))
    return 0


if __name__ == "__main__":
    main()
