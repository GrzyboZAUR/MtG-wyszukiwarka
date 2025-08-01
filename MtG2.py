import pandas as pd
import subprocess
import platform
import os
from tkinter import filedialog


# wybór plików:
plik_talie = filedialog.askopenfilename(title="Wybierz plik .xlsx z talią", filetypes=[("Excel files", "*.xlsx")])
plik_csv = filedialog.askopenfilename(title="Wybierz plik CSV do skanowania", filetypes=[("CSV files", "*.csv")])
# folder z wynikiem:
folder_csv = os.path.dirname(plik_csv)
sciazka_wynik = os.path.join(folder_csv, "wynik.csv")

# 1. Wczytaj plik z taliami
df_talie = pd.read_excel(plik_talie, sheet_name='talie')

# 2. Połącz wszystkie kolumny w jedną listę
lista_serii=[
    df_talie[col].dropna()
    for col in df_talie.columns
    if not df_talie[col].dropna().empty
]
if lista_serii:
    wszystkie_karty = pd.concat(lista_serii)
else:
    wszystkie_karty = pd.Series(dtype=str)

# 3. Oczyść listę
wszystkie_karty = wszystkie_karty.dropna().astype(str).str.strip().str.lower().unique()
zbior_kart = set(wszystkie_karty)

# 4. Wczytaj CSV (zakładam, że separator to TAB '\t')
df_csv = pd.read_csv(plik_csv, sep="\t")

# 5. Oczyść kolumnę Name
df_csv['Name_clean'] = df_csv['Name'].astype(str).str.strip().str.lower()

# 6. Sprawdź, czy karta jest w kolekcji
df_csv['W_Kolekcji'] = df_csv['Name_clean'].apply(lambda x: 1 if x in zbior_kart else 0)

# 7. Posortuj – najpierw te w kolekcji
df_csv.sort_values(by='W_Kolekcji', ascending=False, inplace=True)

# 8. Oczyść ceny
df_csv['Cena_float'] = df_csv['Price'].astype(str).str.replace(',', '.').str.extract(r'([\d\.]+)').astype(float)

# 9. Suma cen i sumy kolekcji
suma_cen = df_csv.loc[df_csv['W_Kolekcji'] == 1, 'Cena_float'].sum()
suma_w_kolekcji = df_csv['W_Kolekcji'].sum()

# 10. Usuń zbędne kolumny
df_csv.drop(columns=['QuantityX', 'Foil', 'Name_clean'], inplace=True, errors='ignore')

# 11. Zmień nazwy kolumn
df_csv.rename(columns={
    'Price': f'Cena_{suma_cen:.2f}',
    'W_Kolekcji': f'W_kolekcji_{int(suma_w_kolekcji)}'
}, inplace=True)

# 12. Usuń tymczasową kolumnę z liczbą
df_csv.drop(columns='Cena_float', inplace=True)

# 13. Zapisz wynik
df_csv.to_csv(sciazka_wynik, index=False)

print(f"✅ Zapisano wynik do: {sciazka_wynik}")

# do ogarnięcia: automatyczne włączanie pliku na każdynm systmie 

def otworz_plik(path):
    system = platform.system()
    if system == 'Windows':
        os.startfile(path)
    elif system == 'Darwin':  # macOS
        subprocess.run(['open', path])
    elif system == 'Linux':
        subprocess.run(['xdg-open', path])
    else:
        print("❌ Nieobsługiwany system operacyjny")

otworz_plik(sciazka_wynik)