import os
import csv

from dotenv import load_dotenv
from imapclient import IMAPClient

load_dotenv()  # take environment variables from .env.


def fetch_mails(user, password, folder):
    """Fetch emails from a specific folder of a user."""
    emails = dict()
    server = IMAPClient(
        host=os.getenv("MAIL_IMAP_HOST"),
        port=os.getenv("MAIL_IMAP_PORT"),
        use_uid=True,
        ssl=bool(os.getenv("MAIL_IMAP_SSL")),
    )
    server.login(user, password)
    server.select_folder(folder)
    emails = server.fetch(server.search(), ["ENVELOPE"])
    server.close_folder()
    server.logout()
    return emails


sent_mails = dict()
sent_mails.update(
    fetch_mails(
        os.getenv("MAIL_USER_1"), os.getenv("MAIL_PASSWORD_1"), "Gesendete Objekte"
    )
)
sent_mails.update(
    fetch_mails(
        os.getenv("MAIL_USER_2"), os.getenv("MAIL_PASSWORD_2"), "Gesendete Objekte"
    )
)

inbox_mails = dict()
inbox_mails.update(
    fetch_mails(os.getenv("MAIL_USER_1"), os.getenv("MAIL_PASSWORD_1"), "INBOX")
)
inbox_mails.update(
    fetch_mails(os.getenv("MAIL_USER_2"), os.getenv("MAIL_PASSWORD_2"), "INBOX")
)

csv_headers = []
csv_rows = [[]]

with open("Test.csv", newline="", encoding="utf-8") as csvfile:
    reader = csv.reader(csvfile, delimiter="\n", quotechar="|")
    for i, row in enumerate(reader):
        cells = row[0].split(",")
        if i == 0:
            csv_headers = cells
            continue
        cells.append("")
        cells.append("")
        count = 0
        for msgid, data in inbox_mails.items():
            envelope = data[b"ENVELOPE"]
            address = (
                envelope.from_[0].mailbox.decode()
                + "@"
                + envelope.from_[0].host.decode()
            )
            if cells[5] == address:
                count += 1
            cells[len(csv_headers)] = count

        count = 0
        for msgid, data in sent_mails.items():
            envelope = data[b"ENVELOPE"]
            for recipient in envelope.to:
                address = recipient.mailbox.decode() + "@" + recipient.host.decode()
                if cells[5] == address:
                    count += 1
                cells[len(csv_headers) + 1] = count

        csv_rows.append(cells)

with open("emails.csv", "w", newline="", encoding="utf-8") as csvfile:
    writer = csv.writer(
        csvfile, delimiter=",", quotechar="|", quoting=csv.QUOTE_MINIMAL
    )
    writer.writerow(csv_headers + ["Inbox", "Sent"])
    writer.writerows(list(set(map(tuple, csv_rows))))
