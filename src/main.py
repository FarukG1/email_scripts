"""Module to fetch email addresses from sent and inbox emails."""

import os
import csv
import configparser
import tkinter as tk
from tkinter import filedialog

from scripts.colors import Colors
from scripts.fetch import Fetch
from scripts.send import Send

tk.Tk().withdraw()


def menu_fetch(config: dict):
    """Fetch email addresses from sent and inbox emails."""
    fetch = Fetch(config)
    users = []
    for index in range(1, int(config["USERS"]["count"]) + 1):
        users.append(
            [
                config["USERS"][f"username{index}"],
                config["USERS"][f"password{index}"],
            ]
        )

    for user in users:
        Colors.print("cyan", f"Fetching emails for {user[0]}...")
        fetch.fetch(user[0], user[1])

    sent_count = len(fetch.sent_addresses)
    inbox_count = len(fetch.inbox_addresses)
    no_reply_count = len(fetch.sent_addresses.difference(fetch.inbox_addresses))

    # Print Statistics
    Colors.print("purple", f"Sent Email Addresses Count: {sent_count}")
    Colors.print("purple", f"Inbox Email Addresses Count: {inbox_count}")
    Colors.print("purple", f"No Reply Email Addresses Count: {no_reply_count}")

    # Save email addresses to file
    selected_file = filedialog.askopenfilename()

    if not selected_file:
        Colors.print("error", "No file selected.")
        return False

    Colors.print(
        "green",
        f'Selected File: "\x1b]8;;file://{selected_file}/\x1b\\{selected_file}\x1b]8;;\x1b\\"',
    )
    filename = fetch.save(selected_file)
    Colors.print(
        "green",
        f'Writing data to "\x1b]8;;file://{filename}/\x1b\\{filename}\x1b]8;;\x1b\\"',
    )

    return True


def menu_send(config: dict):
    """Send emails to email addresses."""
    send = Send(config)
    subject = "Test"
    body = """
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Test</title>
    </head>
    <body>
        <h1>Test</h1>
        <p>This is a test email.</p>
    </body>
</html>
"""

    # Save email addresses to file
    selected_file = filedialog.askopenfilename()

    if not selected_file:
        Colors.print("error", "No file selected.")
        return False

    # Send emails
    with open(selected_file, "r", encoding="utf-8") as file:
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
            msg = send.create(to, subject, body)
            if send.send(to, msg):
                Colors.print("green", f"Email sent to {to}")

                if send.save(msg):
                    Colors.print("green", "Email saved to sent folder")
                else:
                    Colors.print("error", "Failed to save email to sent folder")
                    continue
            else:
                Colors.print("error", f"Failed to send email to {to}")
                continue

    return True


def main():
    """Main function."""

    # Print ASCII Art
    print(
        rf"""
{Colors.purple}$$$$$$$$\                         $$\ $$\        {Colors.warning}$$$$$$\                      $$\             $$\{Colors.reset}
{Colors.purple}$$  _____|                        \__|$$ |      {Colors.warning}$$  __$$\                     \__|            $$ |{Colors.reset}
{Colors.purple}$$ |      $$$$$$\$$$$\   $$$$$$\  $$\ $$ |      {Colors.warning}$$ /  \__| $$$$$$$\  $$$$$$\  $$\  $$$$$$\  $$$$$$\{Colors.reset}
{Colors.purple}$$$$$\    $$  _$$  _$$\  \____$$\ $$ |$$ |      {Colors.warning}\$$$$$$\  $$  _____|$$  __$$\ $$ |$$  __$$\ \_$$  _|{Colors.reset}
{Colors.purple}$$  __|   $$ / $$ / $$ | $$$$$$$ |$$ |$$ |       {Colors.warning}\____$$\ $$ /      $$ |  \__|$$ |$$ /  $$ |  $$ |{Colors.reset}
{Colors.purple}$$ |      $$ | $$ | $$ |$$  __$$ |$$ |$$ |      {Colors.warning}$$\   $$ |$$ |      $$ |      $$ |$$ |  $$ |  $$ |$$\{Colors.reset}
{Colors.purple}$$$$$$$$\ $$ | $$ | $$ |\$$$$$$$ |$$ |$$ |      {Colors.warning}\$$$$$$  |\$$$$$$$\ $$ |      $$ |$$$$$$$  |  \$$$$  |{Colors.reset}
{Colors.purple}\________|\__| \__| \__| \_______|\__|\__|       {Colors.warning}\______/  \_______|\__|      \__|$$  ____/    \____/{Colors.reset}
                                                                                  {Colors.warning}$$ |{Colors.reset}
                                                                                  {Colors.warning}$$ |{Colors.reset}
                                                                                  {Colors.warning}\__|{Colors.reset}""",
    )

    # Generate initial configuration file
    config = configparser.ConfigParser()
    if not os.path.exists("./script.ini"):
        config["SMTP"] = {
            "host": "<host>",
            "port": "<port>",
            "ssl": "<True/False>",
            "sent_folder": "<Sent Folder>",
            "email": "<email>",
            "username": "<username>",
            "password": "<password>",
        }
        config["IMAP"] = {
            "host": "<host>",
            "port": "<port>",
            "ssl": "<True/False>",
            "sent_folder": "<Sent Folder>",
            "inbox_folder": "<Inbox Folder>",
        }
        config["USERS"] = {
            "count": "1",
            "username1": "<username1>",
            "password1": "<password1>",
        }
        with open("./script.ini", "w", encoding="utf-8") as configfile:
            config.write(configfile)
        Colors.print(
            "error",
            "Please fill out the configuration file (script.ini) and run the script again.",
        )
        return 0

    config.read("./script.ini")

    Colors.print("cyan", "What do you want to do?\n")
    Colors.print("cyan", "\t1. Fetch Email Addresses")
    Colors.print("cyan", "\t2. Send Email")
    Colors.print("cyan", "\t3. Exit\n")

    choice = input(Colors.colored("green", "Enter your choice: "))
    print("\n")

    if choice == "1":
        menu_fetch(config)
    elif choice == "2":
        menu_send(config)

    print("\n\n")
    input(Colors.colored("green", "Press Enter to exit..."))
    return 0


if __name__ == "__main__":
    main()
