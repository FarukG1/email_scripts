"""Fetch emails from an IMAP server."""

import csv

from imapclient import IMAPClient


class Fetch:
    """Fetch emails from an IMAP server."""

    def __init__(self, config: dict):
        """Initialize the class."""
        self.host = config["IMAP"]["host"]
        self.port = int(config["IMAP"]["port"])
        self.ssl = bool(config["IMAP"]["ssl"])
        self.sent_folder = config["IMAP"]["sent_folder"]
        self.inbox_folder = config["IMAP"]["inbox_folder"]
        self.sent_addresses = set()
        self.inbox_addresses = set()

    def _sent(self, username: str, password: str):
        """Fetch sent emails."""
        sent_mails = {}
        server = IMAPClient(
            host=self.host,
            port=self.port,
            ssl=self.ssl,
        )
        server.login(username, password)
        server.select_folder(self.sent_folder)
        sent_mails = server.fetch(server.search(), ["ENVELOPE"])
        server.close_folder()
        server.logout()
        return sent_mails

    def _inbox(self, username: str, password: str):
        """Fetch sent emails."""
        inbox_mails = {}
        server = IMAPClient(
            host=self.host,
            port=self.port,
            ssl=self.ssl,
        )
        server.login(username, password)
        server.select_folder(self.inbox_folder)
        inbox_mails = server.fetch(server.search(), ["ENVELOPE"])
        server.close_folder()
        server.logout()
        return inbox_mails

    def fetch(self, user, password):
        """Fetch email addresses from sent and inbox emails."""
        sent_mails = self._sent(user, password)
        inbox_mails = self._inbox(user, password)

        # Get email addresses from sent emails
        for _, data in sent_mails.items():
            envelope = data[b"ENVELOPE"]
            address = (
                envelope.to[0].mailbox.decode(encoding="utf-8")
                + "@"
                + envelope.to[0].host.decode(encoding="utf-8")
            )
            self.sent_addresses.add(address.lower())

        # Get email addresses from inbox emails
        for _, data in inbox_mails.items():
            envelope = data[b"ENVELOPE"]
            address = (
                envelope.from_[0].mailbox.decode(encoding="utf-8")
                + "@"
                + envelope.from_[0].host.decode(encoding="utf-8")
            )
            self.inbox_addresses.add(address.lower())

        return True

    def save(self, file: str, all_emails: bool = False):
        """Save data to a CSV file."""

        csv_header = []
        csv_rows = []
        non_csv_rows = []
        with open(file, "r", encoding="utf-8") as csvfile:
            csv_reader = csv.reader(csvfile)
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

            if all_emails:
                for address in self.sent_addresses:
                    non_csv_rows.append([None] * len(csv_header))
                    non_csv_rows[-1][email_index] = address
            else:
                for address in self.sent_addresses.difference(self.inbox_addresses):
                    non_csv_rows.append([None] * len(csv_header))
                    non_csv_rows[-1][email_index] = address

            for row in csv_reader:
                if row[email_index] in self.sent_addresses.difference(
                    self.inbox_addresses
                ):
                    csv_rows.append(row)
                    array = [None] * len(csv_header)
                    array[email_index] = row[email_index]
                    if array in non_csv_rows:
                        non_csv_rows.remove(array)

        filename = file.split(".csv")[0] + "_no_reply.csv"
        with open(filename, "w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(
                csvfile, delimiter=",", quotechar="|", quoting=csv.QUOTE_MINIMAL
            )
            writer.writerow(csv_header)
            writer.writerows(non_csv_rows)
            writer.writerows(csv_rows)
        return filename
