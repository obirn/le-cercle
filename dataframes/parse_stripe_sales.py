import pandas as pd

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

perfumes_to_short = {
    'VAGUE DE FOLIE VERTE': 'VFL',
    "LA DAME BLANCHE": 'LDB',
    "OSMANTHÉ": 'OSM',
    "L'EAU À LA BOUCHE": 'ELB',
    "LIME ABSOLUE": 'LIM',
    "À L'IRIS": 'IRI',
    "MAGNOL'ART": 'MGA',
}


conditioning_to_short = {
    "Coffret 75 ml": "75C",
    "Coffret 30 ml": "30C",
    "Flacon seul 75 ml": "75S",
    "Flacon seul 30 ml": "30S",
}


def parse_stripe_sales():
    print("Parsing stripe sales... \n")
    df = pd.read_csv(load_path + csv_path + "unified_payments.csv")

    stripe_columns = \
        ['id', 'Description', 'Seller Message', 'Created (UTC)', 'Amount',
         'Amount Refunded', 'Currency', 'Converted Amount',
         'Converted Amount Refunded', 'Fee', 'Tax', 'Converted Currency', 'Mode',
         'Status', 'Statement Descriptor', 'Customer ID', 'Customer Description',
         'Customer Email', 'Captured', 'Card ID', 'Card Last4', 'Card Brand',
         'Card Funding', 'Card Exp Month', 'Card Exp Year', 'Card Name',
         'Card Address Line1', 'Card Address Line2', 'Card Address City',
         'Card Address State', 'Card Address Country', 'Card Address Zip',
         'Card Issue Country', 'Card Fingerprint', 'Card CVC Status',
         'Card AVS Zip Status', 'Card AVS Line1 Status',
         'Card Tokenization Method', 'Shipping Address Line1',
         'Shipping Address Line2', 'Shipping Address City',
         'Shipping Address State', 'Shipping Address Country',
         'Shipping Address Postal Code', 'Disputed Amount', 'Dispute Status',
         'Dispute Reason', 'Dispute Date (UTC)', 'Dispute Evidence Due (UTC)',
         'Invoice ID', 'Invoice Number', 'Payment Source Type', 'Is Link',
         'Destination', 'Transfer', 'Interchange Costs',
         'Merchant Service Charge', 'Transfer Group', 'PaymentIntent ID',
         ]

    wix_columns = ['Customer IP (metadata)', 'wix_transaction_id (metadata)',
                   'Billing Address (metadata)', 'Customer Email (metadata)',
                   'Shipping Address (metadata)', 'Order Item #1 (metadata)',
                   'Customer Phone (metadata)', 'Invoice Id (metadata)',
                   'Customer Name (metadata)', 'Included Charges: Shipping (metadata)',
                   'Order Item #2 (metadata)', 'Order Item #3 (metadata)']

    # Remove 'Test' orders of the dataframe
    print("Removing test orders...")
    print("Number of orders before removing:", len(df))
    df['Created (UTC)'] = pd.to_datetime(df["Created (UTC)"])
    df = df[~(df['Created (UTC)'] <= '13/10/2021  18:45:00')]
    print("Number of orders after removing:", len(df))
    print("\n")

    print("Removing un-paid orders...")
    print("Number of orders before removing:", len(df))
    # Keep orders that have been paid successfully
    df = df.loc[(df["Status"] == "Paid")]
    print("Number of orders after removing:", len(df))
    print("\n")

    print("Parsing orders...")
    df = parse_stripe_orders(df)

    # Reset index
    df = df.reset_index()

    # Save dataframe as excel
    df.to_excel(save_path + csv_path + "stripe_sales_product_as_columns.xlsx")

    return df


def parse_stripe_orders(df: pd.DataFrame):
    ordered_products = pd.DataFrame(
        0, columns=products_as_column + ["Country"], index=df.index)
    for n, row in df.iterrows():
        
        pays = row["Shipping Address (metadata)"]
        ordered_products.at[n, "Country"] = pays[-2::]

        order = row["Description"]
        print(f"Order: {order}")
        products = order.split(";")
        for product in products:
            # print(product)
            column_name = ""
            product = product.strip()
            product_words = product.split(" ")

            # Extract quantity

            quantity = 1
            start_index = product.find('[')
            if start_index != -1:
                end_index = product.find("]")
                quantity = int(product[start_index+1:end_index])
                product = product[:start_index]

            if product in ["ENSEMBLE D'ÉCHANTILLONS", "COFFRET DÉCOUVERTE"]:
                column_name = product
            elif product_words[0] == "ÉCHANTILLON":
                perfume = product[12:].strip()
                column_name = perfumes_to_short[perfume] + " E"
            else:
                [perfume, conditioning] = product.split("Conditionnement:")
                perfume, conditioning = perfume.strip(), conditioning.strip()
                shortName = perfumes_to_short[perfume]
                shortConditioning = conditioning_to_short[conditioning]
                column_name = shortName + " " + shortConditioning
            print(f'   Product: {product}')
            print(f"   Adding {quantity} to {column_name}")
            ordered_products.at[n, column_name] += quantity
            


    df["Source"] = "Wix"
    return df.join(ordered_products)


if __name__ == "__main__":
    parse_stripe_sales()
