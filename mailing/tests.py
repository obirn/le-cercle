import unittest
from mailing_smtp_imap import *
import pandas as pd

# TODO: Add tests for nan values
# TODO: Add tests for double sent clients
# TODO: Add tests for language
# TODO: Add tests for imap and outlook servers

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