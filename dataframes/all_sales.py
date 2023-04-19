from parse_stripe_sales import parse_stripe_sales
from parse_papa_sales import parse_papa_sales
import pandas as pd
import numpy as np


load_path = "../../Data/Load/"
save_path = "../../Data/Save/"
excel_path = "Excels/"
csv_path = "CSVs/"
sales_path = "Ventes/"
clients_path = "Clients/"

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
                                 'Shipping Address (metadata)', 'Included Charges: Shipping (metadata)', 'Source', 'Country'] + products_as_column

    stripe_df = stripe_df[stripe_accounting_columns]

    # Normalize columns names
    stripe_to_normalized = {
        'Created (UTC)': 'Date',
        'Customer Email (metadata)': 'Email',
        'Customer Name (metadata)': 'Nom complet',
        'Customer Phone (metadata)': 'Téléphone',
        'Shipping Address (metadata)': 'Adresse complète',
        'Included Charges: Shipping(metadata)': 'Frais de livraison',
        'Country': "Pays"
    }

    stripe_df.rename(columns=stripe_to_normalized, inplace=True)

    print(stripe_df.loc[stripe_df["Email"] == "client@example.com"]["Pays"])

    
    all_sales = pd.concat([papa_df, stripe_df], join="outer", axis=0)

    # Load wix contacts csv
    wix_contacts = pd.read_csv(load_path + csv_path + "contacts-wix-06-04-23.csv")

    # Rename columns
    wix_contacts = wix_contacts.rename(columns={"E-mail 1": "Email", "Nom de famille": "Nom"})

    sales_clients = all_sales[["Email", "Nom", "Prénom", "Civilité", "Pays"]]
    merged_df = pd.merge(sales_clients, wix_contacts, on='Email', how='left')
    merged_df['Nom'] = merged_df['Nom_y'].fillna(merged_df['Nom_x'])
    merged_df['Prénom'] = merged_df['Prénom_y'].fillna(merged_df['Prénom_x'])
    
    clients_by_email : pd.DataFrame = merged_df.groupby(['Email'])['Nom', 'Prénom', 'Civilité', "Pays"].agg({'Nom':'first', 'Prénom':'first', 'Civilité': 'first', "Pays": "first"}).reset_index()

    clients_by_email.to_excel(save_path + excel_path + clients_path + "clients_by_email.xlsx")

    # Save the dataframe as an excel
    all_sales.to_excel(save_path + excel_path + sales_path + "all_sales_over_time.xlsx")
    return all_sales


def main():
    get_all_sales()


if __name__ == "__main__":
    get_all_sales()
