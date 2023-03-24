import pandas as pd
import numpy as np

excel_path = "../Excels/"
csv_path = "../CSVs/"

perfumes_to_short = {
    'Vague de Folie Verte': 'VFL',
    "La Dame Blanche": 'LDB',
    "Osmanthé": 'OSM',
    "L'eau à la bouche": 'ELB',
    "Lime Absolue": 'LIM',
    "A L'Iris": 'IRI',
    "Magnol'Art": 'MGA'
}

list_of_perfumes = ["Vague de Folie Verte",	"La Dame Blanche", "Osmanthé",
                    "L'eau à la bouche", "Lime Absolue", "A L'Iris", "Magnol'Art"]


def compute_client_data():
    df = pd.read_excel(
        excel_path + "ventes_par_article_26_09_2021__02_03_2023.xlsx", header=0)

    mails = df["E-mail du client"].unique()

    perfumes = df["Nom de l`article"].unique()

    frequently_bought_together = pd.DataFrame(
        np.zeros((len(perfumes), len(perfumes))), index=perfumes[:], columns=perfumes[:])

    perfumes_by_client = pd.DataFrame(
        False, index=mails[:], columns=perfumes[:])

    for mail in mails:
        client_orders = df.loc[df["E-mail du client"] == mail]
        client_known_perfumes = client_orders["Nom de l`article"].unique()
        for i in range(len(client_known_perfumes)):
            curr = client_known_perfumes[i]
            perfumes_by_client[curr][mail] = True

            for j in range(i+1, len(client_known_perfumes)):
                adj = client_known_perfumes[j]

                frequently_bought_together[curr][adj] += 1
                frequently_bought_together[adj][curr] += 1

    frequently_bought_together_percentage = pd.DataFrame(
        np.zeros((len(perfumes), len(perfumes))), index=perfumes[:], columns=perfumes[:])

    for curr in perfumes:
        total = sum(frequently_bought_together[curr])
        num_perfumes = len(perfumes)
        for adj in perfumes:
            frequently_bought_together_percentage[curr][adj] = \
                0 if total == 0 else frequently_bought_together[curr][adj] / total

    # Save as excels
    perfumes_by_client.to_excel(excel_path + "perfumes_by_client.xlsx")
    frequently_bought_together.to_excel(
        excel_path + "frequently_bought_together.xlsx")
    frequently_bought_together_percentage.to_excel(
        excel_path + "frequently_bought_together_percentage.xlsx")


def merge_databases():
    papa_df = pd.read_excel(excel_path + "BDD Papa.xlsx", header=1)
    wix_df = pd.read_csv(csv_path + "contacts.csv")

    merged_columns = ["Prénom",	"Nom de famille",	"Email",	"Téléphone", "Rue",	"État/Région",	"Code postal",
                      "Pays",	"Créé le",	"Nombre total d'achats",	"Dépenses totales", "Date de la dernière activité",	"Langue"]

    products = ["OSM 75S", "OSM 75C", "OSM 30S", "OSM 30C", "OSM E",
                "ELB 75S", "ELB 75C", "ELB 30S", "ELB 30C", "ELB E",
                "IRI 75S", "IRI 75C", "IRI 30S", "IRI 30C", "IRI E",
                "LIM 75S", "LIM 75C", "LIM 30S", "LIM 30C", "LIM E",
                "MGA 75S", "MGA 75C", "MGA 30S", "MGA 30C", "MGA E",
                "VFL 75S", "VFL 75C", "VFL 30S", "VFL 30C", "VFL E",
                "LDB 75S", "LDB 75C", "LDB 30S", "LDB 30C", "LDB E",
                "COFFRET DECOUVERTE"]

    merged_columns += products

    # Compute all unique emails
    papa_unique_emails = papa_df['Email'].drop_duplicates()
    wix_unique_emails = wix_df['E-mail 1'].drop_duplicates()
    unique_emails = pd.concat(
        [papa_unique_emails, wix_unique_emails]).drop_duplicates()

    print("Number of unique emails: ", len(unique_emails))

    merged_df = pd.DataFrame(
        data=None, columns=merged_columns)

    merged_df['Email'] = unique_emails

    print(merged_df)

    merge_papa_df(merged_df, papa_df)


def count_crosses(s):
    if s == 'NaN':
        return 0
    else:
        return s.lower().count('x')

# Define function to merge rows and keep longest strings


# def merge_longest_strings(df: pd.DataFrame):
#     result = pd.Series()
#     for col in df.columns:
#         result[col] = df[col].str.len().idxmax()
#     return df.loc[result]


def merge_papa_df(merged_df: pd.DataFrame, papa_df: pd.DataFrame):
    # Get client infos
    client_infos = papa_df[papa_df.columns[~papa_df.columns.isin(
        ['30 ml', '75 ml', 'Ech', "Vague de Folie Verte",	"La Dame Blanche",
         "Osmanthé",	"L'eau à la bouche",	"Lime Absolue",	"A L'Iris",	"Magnol'Art", 'Objet'])]]

    # Sort the DataFrame by date field in descending order
    client_infos = client_infos.sort_values(by='Date', ascending=False)

    # Drop duplicates based on email field and keep the latest date row
    client_infos = client_infos.drop_duplicates(subset='Email', keep='first')

    # Replace cross by number of crosses in products columns
    for column in list_of_perfumes:
        papa_df[column] = papa_df[column].apply(
            lambda x: count_crosses(str(x)))

    # Split dataframe depending on the format
    papa_ech = papa_df.loc[papa_df['Ech'].notnull()]
    papa_75 = papa_df.loc[papa_df['75 ml'].notnull()]
    papa_30 = papa_df.loc[papa_df['30 ml'].notnull()]

    products_dataframes = [papa_ech, papa_30, papa_75]

    aggregate_dict = dict.fromkeys(list_of_perfumes, 'sum')

    for i in range(len(products_dataframes)):

        # Merge duplicates and sum products
        products_dataframes[i] = products_dataframes[i].groupby(
            "Email").agg(aggregate_dict).reset_index()

        # Rename columns so that it has the same columns as merged_df
        rename_dict = perfumes_to_short.copy()
        if (i == 0):
            variant = 'E'
        elif (i == 1):
            variant = '75C'
        else:
            variant = '30C'

        for key, value in rename_dict.items():
            rename_dict[key] = value + ' ' + variant

        products_dataframes[i].rename(columns=rename_dict, inplace=True)

        commun_columns = products_dataframes[i].columns.intersection(
            merged_df.columns).to_list()[1:]

        for column in commun_columns:
            a = merged_df[column].add(
                products_dataframes[i][column], fill_value=0)
            print(a)
            input()


def main():
    merge_databases()


if __name__ == "__main__":
    main()
