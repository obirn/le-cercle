import pandas as pd

products = ["OSM 75S", "OSM 75C", "OSM 30S", "OSM 30C", "OSM E",
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


def main():
    df = pd.read_excel(
        "C:/Users/robin/Desktop/Perso/Professionel/Le Cercle/Excels/Ventes/ventes-stripe.xlsx")

    # Sort the dataframe by date, and remove 'Test' orders.
    df = df[~(df['Created (UTC)'] < '2018-04-01')]
    orders = df["Description"]

    # Remove last element - first order, as it is an old
    for n, order in orders.items():
        # print(order)
        parse_order(order)
        # input()


def parse_order(order: str):
    products = order.split(";")
    for product in products:
        print(product)
        column_name = ""
        product = product.strip()
        product_words = product.split(" ")

        # Extract quantity
        start_index = product.find('[')
        if start_index != -1:
            end_index = product.find("]")
            quantity = int(product[start_index+1:end_index])
            product = product[:start_index]
            print(quantity)

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

        print(product + "--> " + column_name)
        # input()


if __name__ == "__main__":
    main()
