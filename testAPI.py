from utils.functions import *


def main():
    try:
        consolidate_csv_path = "C:\\Users\\Public\\Downloads\\ACP.RefreshData\\ACP.CM_2025.csv"
        consolidate_parquet_path = "C:\\Users\\Public\\Downloads\\ACP.RefreshData\\ACP.CM_2025.parquet"
        convert_parquet_to_csv(consolidate_parquet_path, consolidate_csv_path)
        #print(f"File converted to CSV: {consolidate_csv_path}")
        # convert_csv_to_parquet(consolidate_csv_path, consolidate_parquet_path)
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
