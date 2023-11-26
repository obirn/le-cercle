import unittest
from mailing_smtp_imap import *
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

class TestStringMethods(unittest.TestCase):
    def test_unsubscribed(self):
        mailing_df : pd.DataFrame = get_mailing_dataframe()
        # Get unsubscribed clients
        unsubscribed_emails : pd.DataFrame = pd.read_excel(
            load_path + excel_dir_path + client_dir_path + "unsubscribed_clients.xlsx")["Email"]
        self.assertFalse(unsubscribed_emails.isin(mailing_df["Email"]).any())

    def test_unique(self):
        mailing_df : pd.DataFrame = get_mailing_dataframe()

        # Get unsubscribed clients
        unsubscribed_emails : pd.DataFrame = pd.read_excel(
            load_path + excel_dir_path + client_dir_path + "unsubscribed_clients.xlsx")["Email"]
        self.assertFalse(unsubscribed_emails.isin(mailing_df["Email"]).any())

if __name__ == '__main__':
    unittest.main()