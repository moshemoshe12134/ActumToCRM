import pandas as pd
from datetime import datetime, timedelta


def create_excel_sheets():
    sheets = ['ProcessedSettlements_Novus', 'ProcessedSettlements_Robin', 'ProcessedReturns_Novus',
              'ProcessedReturns_Robin']
    columns = ['Merchant', 'Funded On', 'Amount', 'Date', 'Due On', 'Paid On', 'Note', 'Total On Date']

    previous_date = (datetime.now() - timedelta(days=1)).strftime("%m_%d_%Y")

    for sheet in sheets:
        df = pd.DataFrame(columns=columns)
        df['Date'] = previous_date  # Add the previous date to the 'Date' column
        filename = f'{sheet}_{previous_date}.xlsx'  # Add the previous date to the filename
        df.to_excel(filename, index=False)


def process_excel_files():
    previous_date = (datetime.now() - timedelta(days=1)).strftime("%m_%d_%Y")

    returns_sheets = [f'Novus Returns_{previous_date}', f'Robin Returns_{previous_date}']
    settlements_sheets = [f'Novus Settlements_{previous_date}', f'Robin Settlements_{previous_date}']

    copy_mapping = {
        f'Novus Returns_{previous_date}': {
            'source_cols': [2, 5, 1, 6],
            'destination_sheet': f'ProcessedReturns_Novus_{previous_date}',
            'destination_cols': [0, 2, 5, 6],
            'exclude_rows': 2
        },
        f'Robin Returns_{previous_date}': {
            'source_cols': [2, 5, 1, 6],
            'destination_sheet': f'ProcessedReturns_Robin_{previous_date}',
            'destination_cols': [0, 2, 5, 6],
            'exclude_rows': 2
        },
        f'Novus Settlements_{previous_date}': {
            'source_cols': [2, 5, 1],
            'destination_sheet': f'ProcessedSettlements_Novus_{previous_date}',
            'destination_cols': [0, 2, 5],
            'exclude_rows': 3
        },
        f'Robin Settlements_{previous_date}': {
            'source_cols': [2, 5, 1],
            'destination_sheet': f'ProcessedSettlements_Robin_{previous_date}',
            'destination_cols': [0, 2, 5],
            'exclude_rows': 3 - 1
        }
    }

    for sheet in returns_sheets + settlements_sheets:
        source_df = pd.read_excel(f'{sheet}.xlsx')
        destination_sheet = copy_mapping[sheet]['destination_sheet']
        destination_cols = copy_mapping[sheet]['destination_cols']
        source_cols = copy_mapping[sheet]['source_cols']
        exclude_rows = copy_mapping[sheet]['exclude_rows'] - 1

        # Exclude rows if specified
        if 'returns' in sheet.lower() and 'novus' in sheet.lower():
            source_df = source_df.iloc[exclude_rows:-1]
        elif 'returns' in sheet.lower() and 'robin' in sheet.lower():
            source_df = source_df.iloc[exclude_rows:-1]
        elif 'settlements' in sheet.lower() and 'novus' in sheet.lower():
            source_df = source_df.iloc[exclude_rows:]
        elif 'settlements' in sheet.lower() and 'robin' in sheet.lower():
            source_df = source_df.iloc[exclude_rows:]

        # Create DataFrame with all columns but no data
        destination_df = pd.DataFrame(columns=range(max(destination_cols) + 1))
        # Fill in the necessary columns from source_df
        for src, dest in zip(source_cols, destination_cols):
            destination_df[dest] = source_df[src]
        destination_df.to_excel(f'{destination_sheet}.xlsx', index=False)


'''
try:
    create_excel_sheets()
    process_excel_files()
except Exception as e:
    print(e)
'''
