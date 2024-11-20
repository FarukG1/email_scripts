import csv
from imapclient import IMAPClient


class Colors:
    """Farben für die Ausgabe im Terminal."""

    PURPLE = "\033[95m"
    BLUE = "\033[94m"
    CYAN = "\033[96m"
    GREEN = "\033[92m"
    WARNING = "\033[93m"
    ERROR = "\033[91m"
    RESET = "\033[0m"
    BOLD = "\033[1m"
    UNDERLINE = "\033[4m"


def colored(color, text):
    """Fügt Farbe zum Text hinzu."""
    return f"{color}{text}{Colors.RESET}"


def fetch_mails(username, password):
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


users = [["kooperation@shinie.de", "SinanEmre9998@"], ["lena@shinie.de", "asdasd1998@"]]

count = 0
sent_addresses = []
inbox_addresses = []
print(colored(Colors.PURPLE, "Starting search..."))
for user in users:
    print(colored(Colors.CYAN, f"Fetching emails for {user[0]}..."))
    sent_mails, inbox_mails = fetch_mails(user[0], user[1])
    with open("data/emails.csv", newline="", encoding="utf-8") as csvfile:
        reader = csv.reader(csvfile, delimiter="\n", quotechar="|")

        for i, row in enumerate(reader):

            for msgid, data in sent_mails.items():
                envelope = data[b"ENVELOPE"]
                address = (
                    envelope.to[0].mailbox.decode(encoding="utf-8")
                    + "@"
                    + envelope.to[0].host.decode(encoding="utf-8")
                )
                if row and row[0] == address:
                    sent_addresses.append(address)

            for msgid, data in inbox_mails.items():
                envelope = data[b"ENVELOPE"]
                address = (
                    envelope.from_[0].mailbox.decode(encoding="utf-8")
                    + "@"
                    + envelope.from_[0].host.decode(encoding="utf-8")
                )
                if row and row[0] == address:
                    inbox_addresses.append(address)

        print(colored(Colors.CYAN, f"Finished fetching emails for {user[0]}."))

# for address in set(sent_addresses):
#     print(colored(Colors.GREEN, "Sent: " + address))

for address in set(inbox_addresses):
    # print(colored(Colors.ERROR, "Inbox: " + address))
    count += 1
print(colored(Colors.ERROR, f"Total Errors: {count}"))

csv_rows = []
for address in set(sent_addresses).difference(set(inbox_addresses)):
    csv_rows.append([address])
with open("data/fixed_emails.csv", "w", newline="", encoding="utf-8") as csvfile:
    writer = csv.writer(
        csvfile, delimiter=",", quotechar="|", quoting=csv.QUOTE_MINIMAL
    )
    writer.writerow(["Email"])
    writer.writerows(csv_rows)
