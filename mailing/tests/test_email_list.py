import sys

# caution: path[0] is reserved for script path (or '' in REPL)
sys.path.append("/home/obirn/Professionnel/Le Cercle/Scripts/mailing/src/")

from src.mailing_smtp_imap import *
import pandas as pd

# TODO: Add tests for nan values
# TODO: Add tests for double sent clients
# TODO: Add tests for language
# TODO: Add tests for subject
# TODO: Add tests for imap and outlook servers
# TODO: Add tests for email - preview (Logo Light....)
# TODO: Add tests to check if the mailing contains french and english mails
# TODO: Check if the email is the correct one (i.e. not sending the black friday email for christmas)
# TODO: Ask the user if he thought to create the discount codes related to the email
# TODO: Check if the links are working
# TODO: Use pytest

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


def test_none_values():
    mailing_df = get_mailing_dataframe()
    email_list = get_email_object_list(mail_dir_path + mail_name, mailing_df)
    assert not None in email_list
