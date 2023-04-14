import pandas as pd

load_path = "../../Data/Load/"
save_path = "../../Data/Save"
excel_path = "Excels/"
csv_path = "CSVs/"
sales_path = "Ventes/"
clients_path = "Clients/"

def keep_last_perfume(row):
    print("row: ", row, "\n")
    return row

def main():
    sales_df = pd.read_excel(load_path + excel_path + sales_path + "all_sales_over_time.xlsx")

    # convert the date column to a datetime object
    sales_df['Date'] = pd.to_datetime(sales_df['Date'])

    # sort the dataframe by email and date
    sales_df = sales_df.sort_values(['Email', 'Date'])
    
    # get the latest purchase for each email
    by_email = sales_df.groupby(["Email"]).last()

    by_email.to_excel(excel_path + sales_path + "client_last_perfumes.xlsx")

if __name__ == "__main__":
    main()