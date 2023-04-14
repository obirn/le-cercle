import pandas as pd
import win32com.client as win32
from string import Template
import os
import re

excel_dir_path = "../../Excels/"
client_dir_path = "Clients/"
sales_dir_path = "Ventes/"

mail_dir_path = "./mails/"


def main():

    # Edit this section
    excel_name = "client_last_perfumes.xlsx"
    mail_name = "offre-vfl/"

    # Get sales dataframe from excel
    sales_df = pd.read_excel(excel_dir_path + sales_dir_path + excel_name)

    # Get unsubscribed clients
    unsubscribed_emails = pd.read_excel(excel_dir_path + client_dir_path + "unsubscribed_clients.xlsx")["Email"]

    # Keep only subscribed clients
    sales_df = sales_df[~sales_df["Email"].isin(unsubscribed_emails)]

    # Confirm the user that he wants to send the mail
    input(f"Press enter if you really want to send the email to {len(sales_df)} people.")

    # Send the mailing
    send_mailing(sales_df, mail_dir_path + mail_name)

def mail_get_template(mail_dir: str, language: str):
    with open(mail_dir+language+".html", encoding='utf-8', mode="r") as file:

        # Replace src by cid in html code 
        file_contents = file.read()
        file_contents = re.sub("src=\"images/", "src=\"cid:", file_contents)
        return Template(file_contents)
    
def add_images_as_attachments(mail: win32.CDispatch, mail_dir: str):
    PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    img_dir = mail_dir + "images/"
    onlyfiles = [f for f in os.listdir(img_dir) if os.path.isfile(os.path.join(img_dir, f))]
    for img_name in onlyfiles:
        absolute_img_path = os.getcwd() + "/" + img_dir + img_name
        attachment = mail.Attachments.Add(absolute_img_path)
        attachment.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, absolute_img_path)


def send_mailing(df: pd.DataFrame, mail_dir_path: str):

    # Load Outlook client
    outlook = win32.Dispatch('outlook.application')
    print("Loading Outlook Client OK")

    # Get sender account
    sender_email = "serviceclient@lecercledesparfumeurscreateurs.com"
    sender_account = None
    for account in outlook.Session.Accounts:
        if account.DisplayName == sender_email:
            sender_account = account
            break
    if sender_account is None:
        print("Error: couldn't find sender account with email "+sender_email)
    print("Getting Sender account OK")

    fr_mail_template = mail_get_template(mail_dir_path, "fr")
    en_mail_template = mail_get_template(mail_dir_path, "en")
    print("Loading emails templates OK")


    fr_mail_subject = "Un Vague De Folie Verte offert !"
    en_mail_subject = "A Vague De Folie Verte for Free !"

    for index, row in df.iterrows():

        receiver_name = row["Prénom"]
        # receiver_email = row["Adresse mail"]
        receiver_email = "robin.varliette@gmail.com"
        isFrench = row["Pays"] in ["France", "Belgique"]

        if isFrench:
            mail_template = fr_mail_template
            mail_subject = fr_mail_subject
        else:
            mail_template = en_mail_template
            mail_subject = en_mail_subject


        # Create mail object
        mail = outlook.CreateItem(0)

        # Add images to the mail
        add_images_as_attachments(mail, mail_dir_path)

        # Attribute the correct sender account
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, sender_account))

        # Set email's Subject
        mail.Subject = mail_subject

        # Format the mail with client name and email
        mail_html_formatted = mail_template.safe_substitute(name=receiver_name)
        mail.HTMLBody = mail_html_formatted
        mail.To = receiver_email

        # Send email
        mail.Send()
        print(f"Mail sent at {receiver_email}")
        input()


if __name__ == "__main__":
    main()
