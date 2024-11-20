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
csv_rows = []

print(colored(Colors.PURPLE, "Starting sorting..."))
for user in users:
    print(colored(Colors.CYAN, f"Fetching emails for {user[0]}..."))

    sent_mails, inbox_mails = fetch_mails(user[0], user[1])

    sent_addresses = []
    inbox_addresses = []

    print(colored(Colors.CYAN, "Fetching sent emails."))
    for msgid, data in sent_mails.items():
        envelope = data[b"ENVELOPE"]
        if envelope.to:
            for addr in envelope.to:
                address = (
                    addr.mailbox.decode(encoding="utf-8")
                    + "@"
                    + addr.host.decode(encoding="utf-8")
                )
                sent_addresses.append(address)

    print(colored(Colors.CYAN, "Fetching inbox emails."))
    for msgid, data in inbox_mails.items():
        envelope = data[b"ENVELOPE"]
        if envelope.from_:
            for addr in envelope.from_:
                address = (
                    addr.mailbox.decode(encoding="utf-8")
                    + "@"
                    + addr.host.decode(encoding="utf-8")
                )
                inbox_addresses.append(address)

    lower_sent_addresses = {addr: addr.lower() for addr in set(sent_addresses)}
    lower_inbox_addresses = {addr: addr.lower() for addr in set(inbox_addresses)}

    no_reply_addresses = set(lower_sent_addresses.values()).difference(
        set(lower_inbox_addresses.values())
    )

    for address in set(lower_sent_addresses.values()).difference(
        set(lower_inbox_addresses.values())
    ):
        original_address = list(lower_sent_addresses.keys())[
            list(lower_sent_addresses.values()).index(address)
        ]
        if [original_address] not in csv_rows:
            csv_rows.append([original_address])

    print(colored(Colors.CYAN, f"Finished fetching emails for {user[0]}."))

with open("data/emails.csv", "w", newline="", encoding="utf-8") as csvfile:
    writer = csv.writer(
        csvfile, delimiter=",", quotechar="|", quoting=csv.QUOTE_MINIMAL
    )
    writer.writerow(["Email"])
    writer.writerows(csv_rows)
