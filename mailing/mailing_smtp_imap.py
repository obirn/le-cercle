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

# The script has to be executed in its directory !
load_path = "../../Data/Load/"
save_path = "../../Data/Save/"
excel_dir_path = "Excels/"
client_dir_path = "Clients/"
sales_dir_path = "Ventes/"
mail_dir_path = "./mails/"
log_dir_path = "./logs/"

def main():

    # Edit this section
    excel_name = "all_sales_over_time.xlsx"
    mail_name = "relance-septembre/"

    # Get sales dataframe from excel
    sales_df = pd.read_excel(
        load_path + excel_dir_path + sales_dir_path + excel_name)

    # Get unsubscribed clients
    unsubscribed_emails = pd.read_excel(
        load_path + excel_dir_path + client_dir_path + "unsubscribed_clients.xlsx")["Email"]

    # Keep only subscribed clients
    sales_df = sales_df[~sales_df["Email"].isin(unsubscribed_emails)]

    # Keep only unique emails
    sales_df = sales_df.drop_duplicates(subset=["Email"])
    
    # Keep only english clients
    # sales_df = sales_df[~sales_df["Pays"].isin(["France", "Belgique", "FR", "BE"])]

    # Define a regular expression pattern for email addresses
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

    # Remove unvalid email adresses
    sales_df = sales_df[sales_df['Email'].str.contains(email_pattern, regex=True)]

    # Confirm the user that he wants to send the mailing
    print(f"This mailing concerns {len(sales_df)} people.")
    print(f"Press [T] to send the mailing to a test email.")
    print(f"Press [Y] to send the mailing to every clients.")
    c = sys.stdin.read(1)
    test_email = None
    if c == "T":
        test_email = input(
            "Enter the email adress the mailing will be sent to (robin.varliette@gmail.com if empty): \n")
        test_email = "robin.varliette@gmail.com" if test_email == "" else test_email
        print(f"[TEST] Sending mailing to {test_email}")
        send_mailing(sales_df, mail_dir_path + mail_name, test_email)
    elif c == "Y":
        confirmation = input("Please confirm by writing \"confirm\": ")
        if confirmation == "confirm":
            print(f"[PROD] Sending mailing to {len(sales_df)} people.")
            send_mailing(sales_df, mail_dir_path + mail_name, test_email)
    else:
        print(f"Unknown option: {c}")


def get_mail_template(mail_path: str, language: str):
    with open(mail_path+language+".html", encoding='utf-8', mode="r") as file:

        # Replace src by cid in html code
        file_contents = file.read()
        file_contents = re.sub("src=\"images/", "src=\"cid:", file_contents)
        with open(mail_path+language+"_save.html", encoding='utf-8', mode="w") as save:
            save.write(file_contents)
        return Template(file_contents)


def add_images_as_attachments(email : EmailMessage, mail_path: str):
    img_dir = mail_path + "images/"
    onlyfiles = [f for f in os.listdir(
        img_dir) if os.path.isfile(os.path.join(img_dir, f))]
    for img_name in onlyfiles:
        absolute_img_path = os.getcwd() + "/" + (img_dir + img_name)
        # know the Content-Type of the image
        maintype, subtype = mimetypes.guess_type(absolute_img_path)[0].split('/')

        img = open(absolute_img_path, "rb")
        # attach it
        email.add_related(img.read(), 
                        maintype=maintype, 
                        subtype=subtype, 
                        cid=img_name)

        img.close()


def send_mailing(df: pd.DataFrame, mail_path: str, test_email: str):

    # Load SMTP client
    print("Loading SMTP Client...")
    smtp_client = smtplib.SMTP_SSL(host='mail.gandi.net',port=465)
    smtp_client.ehlo()
    smtp_client.login(smtp_client_id, smtp_client_pass)
    print("Loading SMTP Server OK")

    # Load IMAP client
    print("Loading IMAP client...")
    imap_client = imaplib.IMAP4_SSL(host="mail.gandi.net",port=993)
    imap_client.login(smtp_client_id, smtp_client_pass)
    print("IMAP client OK")

    fr_mail_template = get_mail_template(mail_path, "fr")
    en_mail_template = get_mail_template(mail_path, "en")
    print("Loading emails templates OK")

    fr_mail_subject = "L'aventure du Cercle des Parfumeurs Créateurs se poursuit !"
    en_mail_subject = "The Adventure of Le Cercle des Parfumeurs Créateurs goes on !"

    last_perfume_bought: pd.DataFrame = pd.read_excel(
        load_path + excel_dir_path + client_dir_path + "client_last_perfumes.xlsx")
    last_perfume_bought = last_perfume_bought.set_index("Email", drop=True)

    # Create a log file
    for _, row in df.iterrows():
        send_mail(row, fr_mail_template, en_mail_template, fr_mail_subject,
                    en_mail_subject, smtp_client, imap_client, mail_path, test_email)

    smtp_client.close()
    imap_client.logout()


def get_most_bought_perfume(email: str, clients_last_perfume: pd.DataFrame):
    most_bought_perfume = str(clients_last_perfume.loc[email].iloc[0])
    product = most_bought_perfume[:3]
    if product == "ENS":
        return "ENSEMBLE D'ECHANTILLONS"
    elif product == "VFL":
        return "Vague de Folie Verte"
    if product == "OSM":
        return "Osmanthé"
    elif product == "ELB":
        return "Eau à la bouche"
    elif product == "LDB":
        return "La Dame Blanche"
    elif product == "IRI":
        return "à l'Iris"
    elif product == "LIM":
        return "Lime Absolue"
    elif product == "MGA":
        return "Magnol'ART"
    else:
        raise Exception(f"unrecognized perfume: {most_bought_perfume[:3]}")


def send_mail(row: pd.Series, fr_mail_template: Template,
              en_mail_template: Template, fr_mail_subject: str,
              en_mail_subject: str, smtp_client, imap_client, mail_path,
              test_email):

    # Load infos (Country, Sex, Name And Family name) of every clients by their emails
    clients_by_email = pd.read_excel(
        load_path + excel_dir_path + client_dir_path + "clients_by_email.xlsx")
    clients_by_email = clients_by_email.set_index("Email", drop=True)

    receiver_email = str(row["Email"])

    client_info = clients_by_email.loc[receiver_email]

    receiver_name = str(client_info["Prénom"]).strip()
    receiver_name = receiver_name.capitalize()
    receiver_name = receiver_name if receiver_name != "Nan" else "Client"
   
    isFrench = client_info["Pays"] in ["France", "Belgique", "FR", "BE"]
    if isFrench:
        mail_template = fr_mail_template
        mail_subject = fr_mail_subject
        if client_info["Civilité"] == "Monsieur":
            greeting = "Cher"
        elif client_info["Civilité"] == "Madame":
            greeting = "Chère"
        else:
            greeting = "Cher(ère)"
    else:
        mail_template = en_mail_template
        mail_subject = en_mail_subject
        greeting = "Dear"

    # Append the name to the greeting.
    greeting += " " + receiver_name + ","

    # Create mail object
    email = EmailMessage()

    # Add images to the mail

    # Attribute the correct sender account
    sender = "serviceclient@lecercledesparfumeurscreateurs.com"
    email["From"] = sender

    # Set email's Subject
    email["Subject"] = mail_subject

    
    # Format the mail
    mail_html_formatted = mail_template.safe_substitute(greeting = greeting);

    email.set_content(mail_html_formatted, subtype="html")

    add_images_as_attachments(email=email, mail_path=mail_path)


    # Send to test e-mail adress if in test mode.
    sendTo = test_email if test_email != None else receiver_email
    email["To"] = sendTo

    # Send email
    try:
        dict_error = "None"
        print("Sending mail to {} for {}".format(sendTo, receiver_email))
        dict_error = smtp_client.sendmail(sender, sendTo, email.as_string())
    except Exception as e:
        print("Couldn't send mail to {}".format(sendTo))
        print("Got exception {}".format(str(e)))
        print("With dictionary: {}".format(str(dict_error)))
    else:
        # print("Mail sent succesfully")
        imap_client.append('Sent', '', imaplib.Time2Internaldate(time()), email.as_string().encode('utf-8'))
        # print("Mail synced with Sent folder on IMAP server successfully")


    if test_email != None:
        input("Press [Enter] to continue")


if __name__ == "__main__":
    main()
