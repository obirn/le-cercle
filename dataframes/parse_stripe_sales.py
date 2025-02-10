import pandas as pd

LOAD_PATH = "../../Data/Load/"
SAVE_PATH = "../../Data/Save/"
EXCEL_PATH = "Excels/"
CSV_PATH = "CSVs/"
SALES_PATH = "Ventes/"
CLIENTS_PATH = "Clients/"

LAST_TEST_ORDER_DATETIME = "13/10/2021  18:45:00"

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

perfumes_to_short = {
    "VAGUE DE FOLIE VERTE": "VFL",
    "LA DAME BLANCHE": "LDB",
    "OSMANTHÉ": "OSM",
    "L'EAU À LA BOUCHE": "ELB",
    "LIME ABSOLUE": "LIM",
    "À L'IRIS": "IRI",
    "MAGNOL'ART": "MGA",
}


conditioning_to_short = {
    "Coffret 75 ml": "75C",
    "Coffret 30 ml": "30C",
    "Flacon seul 75 ml": "75S",
    "Flacon seul 30 ml": "30S",
}


def parse_stripe_sales(stripe_csv_path: str):
    print("Parsing stripe sales: " + stripe_csv_path + "\n")

    # Load csv from path
    df = pd.read_csv(LOAD_PATH + CSV_PATH + stripe_csv_path)

    # Remove 'Test' orders of the dataframe
    df["Created date (UTC)"] = pd.to_datetime(df["Created date (UTC)"])
    df = df[~(df["Created date (UTC)"] <= LAST_TEST_ORDER_DATETIME)]

    # Only keep orders that have been paid successfully
    df = df.loc[(df["Status"] == "Paid")]

    # Parsing orders
    df = parse_stripe_orders(df)

    # Reset index
    df = df.reset_index()

    # Save dataframe as excel
    df.to_excel(SAVE_PATH + CSV_PATH + "stripe_sales_product_as_columns.xlsx")

    return df


def parse_stripe_orders(df: pd.DataFrame):
    # Create a new dataframe
    ordered_products = pd.DataFrame(
        0, columns=products_as_column + ["Country"], index=df.index
    )
    
    # For every order:
    for n, row in df.iterrows():

        # Get country information
        pays = row["Shipping Address (metadata)"]
        ordered_products.at[n, "Country"] = pays[-2::]

        # Get products from order
        order = row["Description"]
        products = order.split(";")

        # For every product
        for product in products:
            product_name = ""
            product = product.strip()
            product_words = product.split(" ")

            # Extract quantity
            quantity = 1
            start_index = product.find("[")
            if start_index != -1:
                end_index = product.find("]")
                quantity = int(product[start_index + 1 : end_index])
                product = product[:start_index]

            if product in ["ENSEMBLE D'ÉCHANTILLONS", "COFFRET DÉCOUVERTE"]:
                product_name = product
            elif product_words[0] == "ÉCHANTILLON":
                perfume = product[12:].strip()
                product_name = perfumes_to_short[perfume] + " E"
            else:
                [perfume, conditioning] = product.split("Conditionnement:")
                perfume, conditioning = perfume.strip(), conditioning.strip()
                shortName = perfumes_to_short[perfume]
                shortConditioning = conditioning_to_short[conditioning]
                product_name = shortName + " " + shortConditioning
            print(f"   Product: {product}")
            print(f"   Adding {quantity} to {product_name}")
            # Add quantity to column name 
            ordered_products.at[n, product_name] += quantity

    df["Source"] = "Wix"
    return df.join(ordered_products)

if __name__ == "__main__":
    parse_stripe_sales()
