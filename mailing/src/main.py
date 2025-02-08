from mailing_smtp_imap import main

# The script has to be executed in the "mailing" directory (can be done use the Makefile) !

# Edit this section
client_excel_filename = "clients_by_email.xlsx"
# client_excel_filename = "test_clients_by_email.xlsx"
# client_excel_filename = "robin_en_test_clients_by_email.xlsx"
client_unsubscribed_filename = "unsubscribed_clients.xlsx"

mail_name = "2025-02-01 saint-valentin/"
en_mail_subject = "Your scented valentine is here!"
fr_mail_subject = "Votre Saint Valentin parfumée est là !"
sender = "serviceclient@lecercledesparfumeurscreateurs.com"
test_mailtrap = False

if __name__ == "__main__":
    main(
        client_excel_filename,
        client_unsubscribed_filename,
        mail_name,
        fr_mail_subject,
        en_mail_subject,
        sender,
        test_mailtrap,
    )