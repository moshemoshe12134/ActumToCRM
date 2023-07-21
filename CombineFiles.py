import pandas as pd
from datetime import datetime, timedelta
import glob


def combine_excel_files():
    previous_date = (datetime.now() - timedelta(days=1)).strftime("%m_%d_%Y")
    file_pattern = f"*{previous_date}.xlsx"
    files = glob.glob(file_pattern)

    combined_data = pd.DataFrame()  # Create an empty DataFrame

    for file in files:
        if 'Processed' in file:
            df = pd.read_excel(file)
            combined_data = pd.concat([combined_data, df], ignore_index=True)

    combined_data.dropna(subset=[combined_data.columns[0]], inplace=True)

    output_file = f"CombinedData_{previous_date}.xlsx"

    # Add the column headers as a new row at the beginning of the DataFrame
    combined_data = pd.concat([pd.DataFrame(['Merchant', 'Funded On', 'Amount', 'Date', 'Due On', 'Paid On', 'Note', 'Total On Date']).transpose(), combined_data], ignore_index=True)

    combined_data.to_excel(output_file, index=False, header=False)

    print(f"Combined file saved at {output_file}")


combine_excel_files()
