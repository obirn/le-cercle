import pandas as pd

from string import Template

from email.message import EmailMessage
from email.utils import make_msgid, formatdate
import mimetypes
import smtplib
import imaplib
from time import time, sleep

import os
import sys
import re

from dotenv import load_dotenv

load_dotenv()

# Environment variables
smtp_client_id = os.getenv("SMTP_CLIENT_ID")
smtp_client_pass = os.getenv("SMTP_CLIENT_PASS")
imap_client_id = os.getenv("IMAP_CLIENT_ID")
imap_client_pass = os.getenv("IMAP_CLIENT_PASS")
mailtrap_smtp_client_id = os.getenv("MAILTRAP_SMTP_CLIENT_ID")
mailtrap_smtp_client_pass = os.getenv("MAILTRAP_SMTP_CLIENT_PASS")

# The script has to be executed in the "mailing" directory (can be done use the Makefile) !
load_path = "../../Data/Load/"
save_path = "../../Data/Save/"
excel_dir_path = "Excels/"
client_dir_path = "Clients/"
sales_dir_path = "Ventes/"
mail_dir_path = "./mails/"
log_dir_path = "./logs/"

sent_log_filename = log_dir_path + "sent.txt"
error_log_filename = log_dir_path + "error.txt"
every_log_filename = log_dir_path + "full_logs.txt"


def init_imap_client():
    imap_client = imaplib.IMAP4_SSL(host="mail.gandi.net", port=993)
    imap_client.login(imap_client_id, imap_client_pass)
    print("IMAP client OK")
    return imap_client


def init_smtp_client(test_mailtrap: bool):
    if test_mailtrap:
        smtp_client = smtplib.SMTP(host="smtp.mailtrap.io", port=2525)
        smtp_client.starttls()
        smtp_client.login(mailtrap_smtp_client_id, mailtrap_smtp_client_pass)
        return smtp_client
    else:
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


def main(
    client_excel_filename,
    client_unsubscribed_filename,
    mail_name,
    fr_mail_subject,
    en_mail_subject,
    sender,
    test_mailtrap,
):
    mail_html_directory = mail_dir_path + mail_name
    mailing_df = get_mailing_dataframe(
        client_excel_filename, client_unsubscribed_filename
    )

    # Confirm the user that he wants to send the mailing
    if test_mailtrap:
        print("..................TESTING MAILTRAP...................")
    else:
        print("..................PRODUCTION MAILING...................")
    print(f"This mailing is named: {mail_name}")
    print(f"The client data has been loaded from the excel: {client_excel_filename}")
    print(f"The french mail subject is: {fr_mail_subject}")
    print(f"The english mail subject is: {en_mail_subject}")
    print(f"This mailing concerns {len(mailing_df)} people.")
    print(".................................................")
    print("")
    print(f"Press [Y] to send the mailing to every clients.")
    c = sys.stdin.read(1)
    if c == "Y":
        if test_mailtrap:
            send_mailing(
                mailing_df,
                mail_html_directory,
                fr_mail_subject,
                en_mail_subject,
                sender,
                test_mailtrap,
            )
            return
        else:
            confirmation = input('Please confirm by writing "confirm": ')
            if confirmation == "confirm":
                print(f"Sending mailing to {len(mailing_df)} people.")
                send_mailing(
                    mailing_df,
                    mail_html_directory,
                    fr_mail_subject,
                    en_mail_subject,
                    sender,
                    test_mailtrap,
                )
            else:
                print('You didn\'t wrote "confirm" correctly.')
    else:
        print(f"Unknown option: {c}")


def get_mailing_dataframe(
    client_excel_filename, client_unsubscribed_filename
) -> pd.DataFrame:
    client_subscribed_path = (
        load_path + excel_dir_path + client_dir_path + client_excel_filename
    )
    client_unsubscribed_path = (
        load_path + excel_dir_path + client_dir_path + client_unsubscribed_filename
    )

    # Get sales dataframe from excel
    mailing_df = pd.read_excel(client_subscribed_path)

    # Get unsubscribed clients
    unsubscribed_emails = pd.read_excel(client_unsubscribed_path)["Email"]

    # Trim every email
    mailing_df["Email"] = mailing_df["Email"].str.strip()
    unsubscribed_emails = unsubscribed_emails.str.strip()

    # Keep only subscribed clients
    mailing_df = mailing_df[~mailing_df["Email"].isin(unsubscribed_emails)]

    # Remove manuel.varliette@free.fr from mailing list
    mailing_df = mailing_df[~mailing_df["Email"].str.contains("manuel.varliette@free.fr")]
    mailing_df = mailing_df[~mailing_df["Email"].str.contains("manuel.varliette@beautyentreprise.com")]

    # Keep only unique emails
    mailing_df = mailing_df.drop_duplicates(subset=["Email"])

    # Define a regular expression pattern for email addresses
    email_pattern = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"

    # Remove unvalid email adresses
    mailing_df = mailing_df[mailing_df["Email"].str.contains(email_pattern, regex=True)]

    return mailing_df


def get_mail_template(mail_html_directory: str, language: str):
    with open(
        mail_html_directory + language + ".html", encoding="utf-8", mode="r"
    ) as file:
        # Replace src by cid in html code
        file_contents = file.read()
        # file_contents = re.sub("images/", "cid:", file_contents)
        with open(
            mail_html_directory + language + "_save.html", encoding="utf-8", mode="w"
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
    mail_html_directory: str,
    fr_mail_subject: str,
    en_mail_subject: str,
    sender: str,
    test_mailtrap: bool,
):
    smtp_client = init_smtp_client(test_mailtrap)
    imap_client = init_imap_client()

    email_list = get_email_object_list(
        mail_html_directory, mailing_df, fr_mail_subject, en_mail_subject, sender
    )

    # Open every log file with append mode
    sent_log_file = open(sent_log_filename, "a")
    error_log_file = open(error_log_filename, "a")
    every_log_file = open(every_log_filename, "a")
    # Write the date of the mailing
    sent_log_file.write("Date: " + str(time()) + "\n")
    error_log_file.write("Date: " + str(time()) + "\n")
    every_log_file.write("Date: " + str(time()) + "\n")

    for email in email_list:
        send_smtp_email(
            imap_client,
            smtp_client,
            email,
            sent_log_file,
            error_log_file,
            every_log_file,
        )
        if test_mailtrap:
            sleep(1)

    # Close every log file
    sent_log_file.close()
    error_log_file.close()
    every_log_file.close()

    smtp_client.close()
    imap_client.logout()


def get_email_object_list(
    mail_path: str,
    mailing_df: pd.DataFrame,
    fr_mail_subject: str,
    en_mail_subject: str,
    sender: str,
) -> list:
    fr_mail_template = get_mail_template(mail_path, "fr")
    en_mail_template = get_mail_template(mail_path, "en")

    print("Emails templates OK")

    email_object_list = []

    for _, client_info in mailing_df.iterrows():
        isFrench = client_info["Pays"] in ["France", "Belgique", "FR", "BE"]
        receiver_name = get_receiver_name(str(client_info["Prénom"]), isFrench)
        receiver_email = str(client_info["Email"]).strip()
        mail_template = fr_mail_template if isFrench else en_mail_template
        subject = fr_mail_subject if isFrench else en_mail_subject
        greeting = get_greeting(isFrench, client_info["Civilité"], receiver_name)

        email_object = create_email_object(
            subject, mail_template, greeting, receiver_email, sender
        )

        # add_images_as_attachments(email_object, mail_path)

        email_object_list.append(email_object)

    return email_object_list


def create_email_object(
    mail_subject: str, mail_template: Template, greeting: str, sendTo: str, sender: str
) -> EmailMessage:
    email = EmailMessage()

    email["From"] = sender
    email["Subject"] = mail_subject
    email["To"] = sendTo
    email["Message-ID"] = make_msgid()
    email["Date"] = formatdate(localtime=True)

    # Format the mail
    mail_html_formatted = mail_template.safe_substitute(greeting=greeting)
    email.set_content(mail_html_formatted, subtype="html")

    return email


def get_receiver_name(receiver_name: str, isFrench: bool):
    receiver_name = receiver_name.strip().capitalize()
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
    sent_log_file,
    error_log_file,
    not_sent_log_file,
):
    try:
        dict_error = "None"
        print("Sending mail to {}".format(email["To"]))
        dict_error = smtp_client.sendmail(email["From"], email["To"], email.as_string())
        sent_log_file.write(email["To"] + "\n")
        sent_log_file.flush()
    except Exception as e:
        error_log_file.write("Couldn't send mail to {}".format(email["To"]))
        error_log_file.write("Got exception {}".format(str(e)))
        error_log_file.write("With dictionary: {}".format(str(dict_error)))
        not_sent_log_file.write(email["To"] + "\n")
        error_log_file.flush()
        not_sent_log_file.flush()
    else:
        # print("Mail sent succesfully")
        imap_client.append(
            "Sent",
            "",
            imaplib.Time2Internaldate(time()),
            email.as_string().encode("utf-8"),
        )
        # print("Mail synced with Sent folder on IMAP server successfully")
