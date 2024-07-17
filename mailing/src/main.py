from mailing_smtp_imap import main

# The script has to be executed in the "mailing" directory (can be done use the Makefile) !

# Edit this section
client_excel_filename = "clients_by_email.xlsx"
client_unsubscribed_filename = "unsubscribed_clients.xlsx"
mail_name = "noel-2023/"
fr_mail_subject = "Ho Ho Ho... Noël avec Le Cercle !"
en_mail_subject = "Ho Ho Ho... Christmas with Le Cercle !"
sender = "serviceclient@lecercledesparfumeurscreateurs.com"

main(
    client_excel_filename,
    client_unsubscribed_filename,
    mail_name,
    fr_mail_subject,
    en_mail_subject,
    sender,
)

if __name__ == "__main__":
    main()
