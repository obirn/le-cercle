import pandas as pd
import win32com.client as win32
from string import Template
import os


def main():
    df = pd.read_csv("./Données clients.csv", sep=';', header=2)
    frenchDf = df.loc[(df["Pays"] == "France") | (df["Pays"] == "Belgique")]
    englishDf = df.loc[(df["Pays"] != "France") & (df["Pays"] != "Belgique")]
    # sendMail(frenchDf, "fr")
    sendMail(englishDf, "en")

def create_mail_template(language: str):
    filename = language + '.html'
    with open(filename, encoding='utf-8', mode="r") as file:
        return Template(file.read())

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

        receiver_name = row["Prénom"]
        receiver_email = row["Adresse mail"]

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
        attachment = mail.Attachments.Add(os.getcwd() + "\\images\\Echantillons_A.jpg")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "Echantillons_A.jpg")

        attachment = mail.Attachments.Add(os.getcwd() + "\\images\\facebook2x.png")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "facebook2x.png")

        attachment = mail.Attachments.Add(os.getcwd() + "\\images\\instagram2x.png")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "instagram2x.png")

        attachment = mail.Attachments.Add(os.getcwd() + "\\images\\LCPC.jpg")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "LCPC.jpg")

        attachment = mail.Attachments.Add(os.getcwd() + "\\images\\leaualabouche_cadre.jpg")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "leaualabouche_cadre.jpg")

        attachment = mail.Attachments.Add(os.getcwd() + "\\images\\logo_transparent.png")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo_transparent.png")

        # Send email
        mail.Send()
        print(f"Mail sent at {receiver_email}")

        input()
    


if __name__ == "__main__":
    main()
