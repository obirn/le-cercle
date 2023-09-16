import pandas as pd
from string import Template

from email.message import EmailMessage
from email.utils import make_msgid
import mimetypes

import smtplib
import os
import sys
import re

load_path = "../../Data/Load/"
save_path = "../../Data/Save/"
excel_dir_path = "Excels/"
client_dir_path = "Clients/"
sales_dir_path = "Ventes/"
mail_dir_path = "./mails/"


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

    # Define a regular expression pattern for email addresses
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

    # Filter the dataframe to keep only rows with valid email addresses
    sales_df = sales_df[sales_df['Email'].str.contains(
        email_pattern, regex=True)]

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

    # Load Smtp server
    print("Loading SMTP Server...")
    smtp_server = smtplib.SMTP_SSL('mail.gandi.net', 465)

    smtp_server.ehlo()

    smtp_server.login("serviceclient@lecercledesparfumeurscreateurs.com", os.getenv("SMTP_CLIENT_PASS"))

    print("Loading SMTP Server OK")

    fr_mail_template = get_mail_template(mail_path, "fr")
    en_mail_template = get_mail_template(mail_path, "en")
    print("Loading emails templates OK")

    fr_mail_subject = "Test 1"
    en_mail_subject = "Test 1"

    last_perfume_bought: pd.DataFrame = pd.read_excel(
        load_path + excel_dir_path + client_dir_path + "client_last_perfumes.xlsx")
    last_perfume_bought = last_perfume_bought.set_index("Email", drop=True)

    for index, row in df.iterrows():
        try:
            send_mail(row, fr_mail_template, en_mail_template, fr_mail_subject,
                    en_mail_subject, smtp_server, mail_path, test_email)
        except Exception as e:
            email = row["Email"]
            print(f"Couldn't send mail to {email}")
            print(e)

    smtp_server.close()    


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
              en_mail_subject: str, smtp_server, mail_path,
              test_email):

    # Load infos (Country, Sex, Name And Family name) of every clients by their emails
    clients_by_email = pd.read_excel(
        load_path + excel_dir_path + client_dir_path + "clients_by_email.xlsx")
    clients_by_email = clients_by_email.set_index("Email", drop=True)

    receiver_email = str(row["Email"])

    client_info = clients_by_email.loc[receiver_email]

    receiver_name = str(client_info["Prénom"]).strip()
    receiver_name = receiver_name if receiver_name != "Nan" else "Client"
    receiver_name = receiver_name.capitalize()

    isFrench = client_info["Pays"] in ["France", "Belgique", "FR", "BE"]
    if isFrench:
        mail_template = fr_mail_template
        mail_subject = fr_mail_subject
        if client_info["Civilité"] == "Monsieur":
            greeting = "Cher "
        elif client_info["Civilité"] == "Madame":
            greeting = "Chère"
        else:
            greeting = "Cher(ère)"
    else:
        mail_template = en_mail_template
        mail_subject = en_mail_subject
        greeting = "Dear"

    # Add the name to the greeting.
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
    smtp_server.sendmail(sender, sendTo, email.as_string())

    print("Mail for {:40s} sent to {}".format(receiver_email, sendTo))

    if test_email != None:
        input("Press [Enter] to continue")


if __name__ == "__main__":
    main()
