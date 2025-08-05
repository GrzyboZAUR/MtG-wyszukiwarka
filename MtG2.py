import pandas as pd
import subprocess
import platform
import os
from tkinter import filedialog


# file selection:
deck_file = filedialog.askopenfilename(title="Select file .xlsx with deck", filetypes=[("Excel files", "*.xlsx")])
csv_file = filedialog.askopenfilename(title="Select CSV file to scan", filetypes=[("CSV files", "*.csv")])
# result folder:
folder_csv = os.path.dirname(csv_file)
path_result = os.path.join(folder_csv, "result.csv")

# 1. Load the file with decks:
df_decks = pd.read_excel(deck_file, sheet_name=0)

# 2. Combine all columns into one list:
series_list=[
    df_decks[col].dropna()
    for col in df_decks.columns
    if not df_decks[col].dropna().empty
]
if series_list:
    all_cards = pd.concat(series_list)
else:
    all_cards = pd.Series(dtype=str)

# 3. Clean up the list:
all_cards = all_cards.dropna().astype(str).str.strip().str.lower().unique()
card_collection = set(all_cards)

# 4. Load the CSV (assuming the separator is TAB '/t')
df_csv = pd.read_csv(csv_file, sep="\t")
if df_csv.iloc[0].astype(str).tolist() == df_csv.columns.astype(str).tolist():
    df_csv = df_csv.iloc[1:]

# 5. Clean up the Name column
df_csv['Name_clean'] = df_csv['Name'].astype(str).str.strip().str.lower()

# 6. Check id the ard is in the collection:
df_csv['In_Collection'] = df_csv['Name_clean'].apply(lambda x: 1 if x in card_collection else 0)

# 7. Sort - those in the collection first:
df_csv.sort_values(by='In_Collection', ascending=False, inplace=True)

# 8. Clean up the prices:
price_column = next((col for col in df_csv.columns if "cena" in col.lower() or "price" in col.lower()), None)
if price_column:
    df_csv["Price_float"] =(
        df_csv[price_column]
        .astype(str)
        .str.replace(',', '.', regex=False)
        .str.extract(r'([\d\.]+)')
        .astype(float)
    )
else:
    df_csv['Price_float'] = 0.0

# 9. Sum the prices and the collection totals:
total_price = df_csv.loc[df_csv['In_Collection'] == 1, 'Price_float'].sum()
total_in_collection = df_csv['In_Collection'].sum()

# 10. Remove unnecessary columns:
df_csv.drop(columns=['QuantityX', 'Foil', 'Name_clean'], inplace=True, errors='ignore')

# 11. Change the name of columns:
df_csv.rename(columns={
    'Price': f'Price_{total_price:.2f}',
    'In_Collection': f'In_Collection_{int(total_in_collection)}'
}, inplace=True)

# 12. Remove the temporary column with the number
df_csv.drop(columns='Price_float', inplace=True)

# 13. Save the result
df_csv.to_csv(path_result, index=False)

print(f"The result was saved to: {path_result}")

# to be understood: automatic inclusion of the file on any system

def open_file(path):
    system = platform.system()
    if system == 'Windows':
        os.startfile(path)
    elif system == 'Darwin':  # macOS
        subprocess.run(['open', path])
    elif system == 'Linux':
        subprocess.run(['xdg-open', path])
    else:
        print("Unsupported operating system")

open_file(path_result)