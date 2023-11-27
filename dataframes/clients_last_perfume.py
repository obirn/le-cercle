import pandas as pd

load_path = "../../Data/Load/"
save_path = "../../Data/Save/"
excel_path = "Excels/"
csv_path = "CSVs/"
sales_path = "Ventes/"
clients_path = "Clients/"

products_as_column = [
    "OSM 75S",
    "OSM 75C",
    "OSM 30S",
    "OSM 30C",
    "OSM E",
    "ELB 75S",
    "ELB 75C",
    "ELB 30S",
    "ELB 30C",
    "ELB E",
    "IRI 75S",
    "IRI 75C",
    "IRI 30S",
    "IRI 30C",
    "IRI E",
    "LIM 75S",
    "LIM 75C",
    "LIM 30S",
    "LIM 30C",
    "LIM E",
    "MGA 75S",
    "MGA 75C",
    "MGA 30S",
    "MGA 30C",
    "MGA E",
    "VFL 75S",
    "VFL 75C",
    "VFL 30S",
    "VFL 30C",
    "VFL E",
    "LDB 75S",
    "LDB 75C",
    "LDB 30S",
    "LDB 30C",
    "LDB E",
    "ENSEMBLE D'ÉCHANTILLONS",
    "COFFRET DÉCOUVERTE",
]


def get_most_bought_perfume(group):
    email = group["Email"]
    perfume_columns = [
        c
        for c in group.columns
        if c.endswith(("S", "C")) and c != "ENSEMBLE D'ÉCHANTILLONS"
    ]
    perfume_sales: pd.Series = group[perfume_columns + ["Date"]].melt(
        id_vars=["Date"], var_name="perfume", value_name="sales"
    )
    if perfume_sales["sales"].sum() > 0:
        most_bought = perfume_sales.sort_values("sales", ascending=False).iloc[0]
        return most_bought["perfume"]
    else:
        #  print(f"no perfume sale for {email}")
        sample_columns = [
            c for c in group.columns if c.endswith("E") and c != "COFFRET DÉCOUVERTE"
        ]
        sample_sales = group[sample_columns + ["Date"]].melt(
            id_vars=["Date"], var_name="sample", value_name="sales"
        )
        if sample_sales["sales"].sum() > 0:
            most_bought = sample_sales.sort_values("sales", ascending=False).iloc[0]
            return most_bought["sample"]
        else:
            others = group[
                ["ENSEMBLE D'ÉCHANTILLONS", "COFFRET DÉCOUVERTE", "Date"]
            ].melt(id_vars=["Date"], var_name="product", value_name="sales")
            # print(others)
            if others["sales"].sum() < 0:
                print(f"no other sale for {email}")
                return "Unknown"
            most_bought = others.sort_values("sales", ascending=False).iloc[0]
            return most_bought["product"]


def main():
    sales_df = pd.read_excel(
        load_path + excel_path + sales_path + "all_sales_over_time.xlsx"
    )

    # convert the date column to a datetime object
    sales_df["Date"] = pd.to_datetime(sales_df["Date"])

    # sort the dataframe by email and date
    sales_df = sales_df.sort_values(["Email", "Date"])

    # get the latest purchase for each email
    by_email = sales_df.groupby(["Email"]).apply(get_most_bought_perfume)

    by_email.to_excel(save_path + excel_path + sales_path + "client_last_perfumes.xlsx")


if __name__ == "__main__":
    main()
