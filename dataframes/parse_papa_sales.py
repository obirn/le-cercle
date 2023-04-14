import pandas as pd

load_path = "../../Data/Load/"
save_path = "../../Data/Save"
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

format_to_short = {
    "30 ml": "30C",
    "75 ml": "75C",
    "Ech": "E",
}


perfumes_to_short = {
    'Vague de Folie Verte': 'VFL',
    "La Dame Blanche": 'LDB',
    "Osmanthé": 'OSM',
    "L'eau à la bouche": 'ELB',
    "Lime Absolue": 'LIM',
    "A L'Iris": 'IRI',
    "Magnol'Art": 'MGA'
}


def get_perfumes_format(row: pd.Series):
    is30 = str(row['30 ml']) != 'nan'
    is75 = str(row['75 ml']) != 'nan'
    isEch = str(row['Ech']) != 'nan'

    if is30:
        if is75 or isEch:
            raise Exception("More than 1 format given")
        return "30C"
    elif is75:
        if is30 or isEch:
            raise Exception("More than 1 format given")
        return "75C"
    elif isEch:
        if is75 or is30:
            raise Exception("More than 1 format given")
        return "E"
    else:
        raise Exception("No format given")


def parse_papa_sales():
    """
        This function parse the 'BDD ventes papa' database and 
        gives a dataframe containing as rows orders, and for columns clients infos + single product quantities.
    """
    df = pd.read_excel(load_path + excel_path + sales_path + "BDD Ventes papa.xlsx", header=1)

    ordered_products = pd.DataFrame(
        0, columns=products_as_column, index=df.index)

    formats = ["30 ml", "75 ml", "Ech"]
    perfumes = ["Vague de Folie Verte",	"La Dame Blanche",
                "Osmanthé",	"L'eau à la bouche",	"Lime Absolue",	"A L'Iris",	"Magnol'Art"]
    for index, row in df.iterrows():
        format = get_perfumes_format(row)

        for perfume in perfumes:
            perfume_quantity = row[perfume]
            try:
                quantity = int(perfume_quantity)
            except:
                quantity = str(perfume_quantity).lower().count('x')

            short_name = perfumes_to_short[perfume]

            column_name = f"{short_name} {format}"

            ordered_products.at[index, column_name] += quantity

    df = df.join(ordered_products)
    df = df.drop(formats + perfumes, axis=1)

    return df


if __name__ == "__main__":
    parse_papa_sales()
