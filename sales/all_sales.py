from parse_stripe_sales import parse_stripe_sales
from parse_papa_sales import parse_papa_sales
import pandas as pd
import numpy as np

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

    stripe_account_columns = ['Created (UTC)', 'Customer Email (metadata)', 'Customer Name (metadata)', 'Customer Phone (metadata)',
                              'Shipping Address (metadata)', 'Included Charges: Shipping (metadata)', 'Source'] + products_as_column

    stripe_df = stripe_df[stripe_account_columns]
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

    all_sales = all_sales.reset_index(drop=True)

    all_sales.to_excel('save.xlsx')
    print(all_sales)

    return


if __name__ == "__main__":
    get_all_sales()
