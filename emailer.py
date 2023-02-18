#!/usr/bin/env python3
"""Bulk email merge tool."""
import argparse
import csv
import email.message
import os
import smtplib
from pathlib import Path

import jinja2
import pycmarkgfm
from bs4 import BeautifulSoup
from bs4 import Comment


def get_data(
    merge_file: Path,
) -> list[dict]:
    """Build list of dicts from CSV file.

    Args:
        merge_file: CSV file path with values to merge into message

    Returns:
        List of dicts
    """
    with merge_file.open("r", encoding="utf-8", newline="") as infile:
        reader = csv.DictReader(infile)
        groupby = "Group"
        if reader.fieldnames and groupby not in reader.fieldnames:
            groupby = "EmailAddress"
        groups = {}
        consolidated = []
        for row in reader:
            members = groups.setdefault(row[groupby], [])
            members.append(row)
        for group in groups.values():
            members = {}
            for member in group:
                for k, v in member.items():
                    children = members.setdefault(k, [])
                    children.append(v)
            if groupby.casefold() == "group":
                del members[groupby]
            consolidated.append(members)
    return consolidated


def send_batch(
    data: list,
    template: jinja2.Template,
    username: str,
    password: str,
    sender: str | None,
) -> None:
    """Send multiple emails.

    Args:
        data: list of dicts for merging in
        template: email message Jinja2 template
        username: mail server username
        password: mail server password
        sender: optional "From" email address
    """
    if sender is None:
        sender = username

    with smtplib.SMTP("smtp.gmail.com", port=587) as server:
        server.starttls()
        server.login(username, password)
        for row in data:
            body = template.render(**row)
            send(server, sender, row["EmailAddress"], body)


def send(server: smtplib.SMTP, sender: str, recipients: list[str], body: str) -> None:
    """Send single email.

    Args:
        server: smtplib SMTP server instance
        sender: "From" email address
        recipients: list of "To" email addresses
        body: email message body
    """
    msg = email.message.EmailMessage()
    msg["From"] = sender
    msg["To"] = recipients
    html_body = pycmarkgfm.gfm_to_html(body, options=pycmarkgfm.options.unsafe)
    soup = BeautifulSoup(html_body, "lxml")
    subject_comment = soup.find(
        string=lambda x: isinstance(x, Comment)
        and x.strip().casefold().startswith("subject")
    )
    if subject_comment:
        msg["Subject"] = str(subject_comment).split(":", 1)[-1].strip()
        subject_comment.extract()
    else:
        msg["Subject"] = ""
    text_body = "".join(soup.find_all(string=True))
    msg.set_content(text_body)
    msg.add_alternative(html_body, subtype="html")
    server.send_message(msg)


def get_template(
    msgfile: Path,
) -> jinja2.Template:
    """Parse message template.

    Args:
        msgfile: path to email message body markdown file

    Returns:
        Message string
    """
    env = jinja2.Environment(
        loader=jinja2.FileSystemLoader(msgfile.parent), autoescape=True
    )
    template = env.get_template(msgfile.name)
    return template


def run():
    """Command runner."""
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "msgfile", help="Path to email message body markdown file", type=Path
    )
    parser.add_argument("csvfile", help="Path to CSV file for variables", type=Path)
    parser.add_argument("-s", "--sender", help="email sender")
    args = parser.parse_args()
    password = os.getenv("SMTPPASSWD", "")
    username = os.getenv("SMTPUSER", "")
    template = get_template(args.msgfile)
    data = get_data(args.csvfile)
    send_batch(data, template, username, password, args.sender)


if __name__ == "__main__":
    run()
