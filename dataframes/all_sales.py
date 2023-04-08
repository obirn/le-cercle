from parse_stripe_sales import parse_stripe_sales
from parse_papa_sales import parse_papa_sales
import pandas as pd
import numpy as np

excel_path = "../../Excels/"
sales_path = "Ventes/"

products_as_column = \
    ["OSM 75S", "OSM 75C", "OSM 30S", "OSM 30C", "OSM E",
     "ELB 75S", "ELB 75C", "ELB 30S", "ELB 30C", "ELB E",
     "IRI 75S", "IRI 75C", "IRI 30S", "IRI 30C", "IRI E",
     "LIM 75S", "LIM 75C", "LIM 30S", "LIM 30C", "LIM E",
     "MGA 75S", "MGA 75C", "MGA 30S", "MGA 30C", "MGA E",
     "VFL 75S", "VFL 75C", "VFL 30S", "VFL 30C", "VFL E",
     "LDB 75S", "LDB 75C", "LDB 30S", "LDB 30C", "LDB E",
     "ENSEMBLE D'ÉCHANTILLONS", "COFFRET DÉCOUVERTE"]


def get_all_sales():
    papa_df = parse_papa_sales()
    stripe_df = parse_stripe_sales()

    stripe_accounting_columns = ['Created (UTC)', 'Customer Email (metadata)', 'Customer Name (metadata)', 'Customer Phone (metadata)',
                                 'Shipping Address (metadata)', 'Included Charges: Shipping (metadata)', 'Source'] + products_as_column

    stripe_df = stripe_df[stripe_accounting_columns]

    # Normalize columns names
    stripe_to_normalized = {
        'Created (UTC)': 'Date',
        'Customer Email (metadata)': 'Email',
        'Customer Name (metadata)': 'Nom complet',
        'Customer Phone (metadata)': 'Téléphone',
        'Shipping Address (metadata)': 'Adresse complète',
        'Included Charges: Shipping(metadata)': 'Frais de livraison',
    }

    stripe_df.rename(columns=stripe_to_normalized, inplace=True)

    all_sales = pd.concat([papa_df, stripe_df], join="outer", axis=0)

    # all_sales = all_sales.reset_index(drop=True)

    # Load wix contacts csv
    wix_contacts = pd.read_csv("../../CSVs/contacts-wix-06-04-23.csv")

    # Set email as index of the dataframe
    wix_contacts.set_index("E-mail 1")

    all_sales["Prénom"] = all_sales.apply(get_name(wix_contacts))

    print(all_sales[all_sales["Prénom"].isna()])

    # Save the dataframe as an excel
    all_sales.to_excel(excel_path + sales_path + "all_sales_over_time.xlsx")
    return all_sales


def get_name(wix_contacts: pd.DataFrame, row: pd.Series, ):
    if not pd.isna(row["Prénom"]):
        return row["Prénom"]
    else:
        return wix_contacts[row["Email"]]["Prénom"]


def main():
    get_all_sales()


if __name__ == "__main__":
    get_all_sales()
