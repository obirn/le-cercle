import pandas as pd
from string import Template
from credentials import *

from email.message import EmailMessage
import mimetypes
import smtplib
import imaplib
from time import time

import os
import sys
import re

# The script has to be executed in the "mailing" directory (can be done use the Makefile) !
load_path = "../../Data/Load/"
save_path = "../../Data/Save/"
excel_dir_path = "Excels/"
client_dir_path = "Clients/"
sales_dir_path = "Ventes/"
mail_dir_path = "./mails/"
log_dir_path = "./logs/"

# Edit this section
client_excel_path = "clients_by_email.xlsx"
mail_name = "noel-2023/"
fr_mail_subject = "Ho Ho Ho... Noël avec Le Cercle !"
en_mail_subject = "Ho Ho Ho... Christmas with Le Cercle !"
sender = "serviceclient@lecercledesparfumeurscreateurs.com"


def init_imap_client():
    imap_client = imaplib.IMAP4_SSL(host="mail.gandi.net", port=993)
    imap_client.login(smtp_client_id, smtp_client_pass)
    print("IMAP client OK")
    return imap_client


def init_smtp_client():
    smtp_client = smtplib.SMTP_SSL(host="mail.gandi.net", port=465)
    smtp_client.ehlo()
    smtp_client.login(smtp_client_id, smtp_client_pass)
    return smtp_client


def send_french_email(test_email: str):
    smtp_client = init_smtp_client()
    imap_client = init_imap_client()
    mailing_df = get_mailing_dataframe()
    mailing_df = mailing_df[mailing_df["Pays"].isin(["France", "Belgique", "FR", "BE"])]
    first_row = mailing_df.iloc[0]
    receiver_email = first_row["Email"]


def main():
    mailing_df = get_mailing_dataframe()
    print(os.getcwd())
    clients_by_email = pd.read_excel(
        load_path + excel_dir_path + client_dir_path + "clients_by_email.xlsx"
    )
    clients_by_email = clients_by_email.set_index("Email", drop=True)

    # Confirm the user that he wants to send the mailing
    print("..................MAILING INFO...................")
    print(f"This mailing is named: {mail_name}")
    print(f"The client data has been loaded from the excel: {client_excel_path}")
    print(f"The french mail subject is: {fr_mail_subject}")
    print(f"The enlgish mail subject is: {en_mail_subject}")
    print(f"This mailing concerns {len(mailing_df)} people.")
    print(".................................................")
    print("")
    print(f"Press [Y] to send the mailing to every clients.")
    c = sys.stdin.read(1)
    test_email = None
    if c == "Y":
        confirmation = input('Please confirm by writing "confirm": ')
        if confirmation == "confirm":
            print(f"Sending mailing to {len(mailing_df)} people.")
            send_mailing(mailing_df, mail_dir_path + mail_name)
        else:
            print('You didn\'t wrote "confirm" correctly.')
    else:
        print(f"Unknown option: {c}")


def get_mailing_dataframe() -> pd.DataFrame:
    # Get sales dataframe from excel
    mailing_df = pd.read_excel(
        load_path + excel_dir_path + client_dir_path + client_excel_path
    )

    # Get unsubscribed clients
    unsubscribed_emails = pd.read_excel(
        load_path + excel_dir_path + client_dir_path + "unsubscribed_clients.xlsx"
    )["Email"]

    # Keep only subscribed clients
    mailing_df = mailing_df[~mailing_df["Email"].isin(unsubscribed_emails)]

    # Keep only unique emails
    mailing_df = mailing_df.drop_duplicates(subset=["Email"])

    # Define a regular expression pattern for email addresses
    email_pattern = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"

    # Remove unvalid email adresses
    mailing_df = mailing_df[mailing_df["Email"].str.contains(email_pattern, regex=True)]

    return mailing_df


def get_mail_template(mail_path: str, language: str):
    print(os.getcwd())
    print(mail_path)
    with open(mail_path + language + ".html", encoding="utf-8", mode="r") as file:
        # Replace src by cid in html code
        file_contents = file.read()
        # file_contents = re.sub("images/", "cid:", file_contents)
        with open(
            mail_path + language + "_save.html", encoding="utf-8", mode="w"
        ) as save:
            save.write(file_contents)
        return Template(file_contents)


def add_images_as_attachments(email: EmailMessage, mail_path: str):
    img_dir = mail_path + "images/"
    onlyfiles = [
        f for f in os.listdir(img_dir) if os.path.isfile(os.path.join(img_dir, f))
    ]
    for img_name in onlyfiles:
        absolute_img_path = os.getcwd() + "/" + (img_dir + img_name)
        # know the Content-Type of the image
        maintype, subtype = mimetypes.guess_type(absolute_img_path)[0].split("/")

        # attach it
        img = open(absolute_img_path, "rb")
        email.add_related(img.read(), maintype=maintype, subtype=subtype, cid=img_name)

        img.close()


def send_mailing(
    mailing_df: pd.DataFrame,
    mail_path: str,
):
    smtp_client = init_smtp_client()
    imap_client = init_imap_client()

    email_list = get_email_object_list(mail_path, mailing_df)

    # for email in email_list:
    #     send_smtp_email(imap_client, smtp_client, email)

    smtp_client.close()
    imap_client.logout()


def get_email_object_list(mail_path: str, mailing_df: pd.DataFrame) -> list:
    fr_mail_template = get_mail_template(mail_path, "fr")
    en_mail_template = get_mail_template(mail_path, "en")

    print("Emails templates OK")

    email_object_list = []

    for _, client_info in mailing_df.iterrows():
        receiver_name = str(client_info["Prénom"]).strip()
        receiver_email = str(client_info["Email"]).strip()
        isFrench = client_info["Pays"] in ["France", "Belgique", "FR", "BE"]
        mail_template = fr_mail_template if isFrench else en_mail_template
        subject = fr_mail_subject if isFrench else en_mail_subject
        greeting = get_greeting(isFrench, client_info["Civilité"], receiver_name)

        email_object = create_email_object(
            subject, mail_template, greeting, receiver_email
        )

        email_object_list.append(email_object)

    return email_object_list


def create_email_object(
    mail_subject: str, mail_template: Template, greeting: str, sendTo: str
) -> EmailMessage:
    email = EmailMessage()

    email["From"] = sender
    email["Subject"] = mail_subject
    email["To"] = sendTo

    # Format the mail
    mail_html_formatted = mail_template.safe_substitute(greeting=greeting)
    email.set_content(mail_html_formatted, subtype="html")

    return email


def get_receiver_name(receiver_name: str, isFrench: bool):
    receiver_name = receiver_name.capitalize()
    if receiver_name != "Nan":
        return receiver_name
    else:
        return "Client" if isFrench else "Customer"


def get_greeting(isFrench: bool, civilité: str, receiver_name: str):
    if isFrench:
        if civilité == "Monsieur":
            greeting = "Cher"
        elif civilité == "Madame":
            greeting = "Chère"
        else:
            greeting = "Cher(ère)"
    else:
        greeting = "Dear"

    greeting += " " + receiver_name + ","
    return greeting


def send_smtp_email(
    imap_client: imaplib.IMAP4_SSL,
    smtp_client: smtplib.SMTP_SSL,
    email: EmailMessage,
):
    try:
        dict_error = "None"
        print("Sending mail to {} for {}".format(email["To"]))
        dict_error = smtp_client.sendmail(email["From"], email["To"], email.as_string())
    except Exception as e:
        print("Couldn't send mail to {}".format(email["To"]))
        print("Got exception {}".format(str(e)))
        print("With dictionary: {}".format(str(dict_error)))
    else:
        # print("Mail sent succesfully")
        imap_client.append(
            "Sent",
            "",
            imaplib.Time2Internaldate(time()),
            email.as_string().encode("utf-8"),
        )
        # print("Mail synced with Sent folder on IMAP server successfully")


if __name__ == "__main__":
    main()
