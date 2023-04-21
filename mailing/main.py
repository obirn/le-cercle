import pandas as pd
import win32com.client as win32
from string import Template
import os

excel_path = "../Excels/"
img_path = "./images/"
mail_path = "./mails"


def main():
    df = pd.read_excel("../../Data/Load/Excels/Clients/client_last_perfumes.xlsx")
    # frenchDf = df.loc[(df["Pays"] == "France") | (df["Pays"] == "Belgique")]
    # englishDf = df.loc[(df["Pays"] != "France") & (df["Pays"] != "Belgique")]
    sendMail(df, "fr")
    # sendMail(englishDf, "en")


def create_mail_template(language: str):
    filename = "mails/saint-valentin/" + language + '.html'
    with open(filename, encoding='utf-8', mode="r") as file:
        return Template(file.read())


def add_images_as_attachments(mail: win32.CDispatch, mail_path: str):
    PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    img_dir = mail_path + "images/"
    onlyfiles = [f for f in os.listdir(img_dir) if os.path.isfile(os.path.join(img_dir, f))]
    for img_name in onlyfiles:
        absolute_img_path = os.getcwd() + "\\" + (img_dir + img_name).replace("/","\\")
        print(absolute_img_path)
        attachment = mail.Attachments.Add(absolute_img_path)
        attachment.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, "images/" + img_name)


def sendMail(df: pd.DataFrame, language: str):

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

    # Get email template based on language
    mail_template = create_mail_template(language=language)

    for index, row in df.iterrows():

        receiver_name = "Robin"
        receiver_email = row["Email"]

        # Create mail object
        mail = outlook.CreateItem(0)

        # Attribute the correct sender account
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, sender_account))

        # Set email's object
        mail.Subject = 'Une Saint-Valentin parfumée ?'

        # Format the mail with client name and email
        mail_with_name = mail_template.safe_substitute(name=receiver_name)
        mail.HTMLBody = mail_with_name
        mail.To = "robin.varliette@gmail.com"

        # Add images to the mail
        add_images_as_attachments(mail, "./mails/saint-valentin/")

        # Send email
        mail.Send()
        input()
        print(f"Mail sent at {receiver_email}")


if __name__ == "__main__":
    main()
