import sys

# caution: path[0] is reserved for script path (or '' in REPL)
# sys.path.append("/home/obirn/Professionnel/Le Cercle/Scripts/mailing/src/")

from src.mailing_smtp_imap import *
import pandas as pd

# TODO: Improve tests for nan values
# TODO: Add tests for language
# TODO: Add tests for subject
# TODO: Add tests for imap and outlook servers
# TODO: Add tests for email - preview (Logo Light....)
# TODO: Add tests to check if the mailing contains french and english mails
# TODO: Check if the email is the correct one (i.e. not sending the black friday email for christmas)
# TODO: Ask the user if he thought to create the discount codes related to the email
# TODO: Check if the links are working
# TODO: Add warning if after removing unsusbriced clients, the mailing list is empty or same length

# Edit this section
client_excel_filename = "clients_by_email.xlsx"
client_unsubscribed_filename = "unsubscribed_clients.xlsx"
mail_name = "noel-2023/"
fr_mail_subject = "Ho Ho Ho... Noël avec Le Cercle !"
en_mail_subject = "Ho Ho Ho... Christmas with Le Cercle !"
sender = "serviceclient@lecercledesparfumeurscreateurs.com"

mail_dir_path = "./mails/"

mailing_df = get_mailing_dataframe(client_excel_filename, client_unsubscribed_filename)
email_list = get_email_object_list(
    mail_dir_path + mail_name, mailing_df, fr_mail_subject, en_mail_subject, sender
)


def get_html_content(email_message):
    # Parcourir les différentes parties du message
    if email_message.is_multipart():
        for part in email_message.walk():
            content_type = part.get_content_type()
            # Vérifier si c'est du HTML
            if content_type == "text/html":
                return part.get_payload(decode=True).decode("utf-8")
    else:
        # Si le message n'est pas multipart, vérifier directement
        if email_message.get_content_type() == "text/html":
            return email_message.get_payload(decode=True).decode("utf-8")

    return None


# Test that the email list does not contain an email that contains "nan" or "Nan" value in his body
def test_nan_values():
    for email in email_list:
        html_content = get_html_content(email)
        # if " nan" in html_content.lower():
        #     print(email["To"])
        #     print(email["From"])
        #     print(html_content)
        assert not " nan," in html_content.lower()


def test_none_values():
    assert not None in email_list


def test_equal_length():
    assert len(mailing_df) == len(email_list)


def test_length_not_zero():
    assert len(email_list) > 0


def test_no_empty():
    for email in email_list:
        assert len(email["To"]) > 0
        assert len(email["From"]) > 0
        assert len(email.get_body()) > 0


def test_double_sent():
    receiver_emails = [email["To"] for email in email_list]
    # Print duplicated in receiver emails
    if len(receiver_emails) != len(set(receiver_emails)):
        print("Duplicated emails:")
        for email in set(receiver_emails):
            if receiver_emails.count(email) > 1:
                print(email)
    assert len(receiver_emails) == len(set(receiver_emails))
