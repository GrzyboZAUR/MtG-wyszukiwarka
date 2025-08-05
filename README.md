A simple Python script that compares scanned Magic: The Gathering cards (CSV format) with a reference deck list (Excel format), and highlights which cards are already in the collection.

---

## 📄 Features

- Loads card data from a CSV file (e.g. exported from apps like [Delver Lens](https://delverlab.com/)).
- Compares it with a deck list stored in an Excel sheet.
- Detects which cards are in the collection.
- Calculates total value of matched cards.
- Saves the result as a new CSV file, sorted by match.
- Opens the result file automatically on Windows.

## 🗂️ File Structure

```bash
.
├── MtG2.py              # Main script
├── requirements.txt     # Dependencies
└── data/
    ├── sample_data.csv         # Example scanned cards (tab-separated)
    └── sample_deck.xlsx        # Example deck list
```
How to Use
1. Make sure you have Python 3.x installed.
2. Install required libraries:

```bash
pip install -r requirements.txt
```

3. Replace the sample files in the data/ folder with your own: 


   - Your scanned cards CSV file (tab-delimited).
   - Your deck list Excel file (make sure the sheet is first).

4. Run the script:

```
python MtG2.py
```
5. The result will be saved in the same folder and opened automatically.

The script creates a new CSV file containing:

 - Whether the card is in your collection.
 - Price (if available).
 - Total value of matched cards.

Requirements
   - pandas
   - openpyxl (for reading .xlsx)
   - platform and os (standard libraries)

Install with:
```
pip install pandas openpyxl

```

License
This project is free to use under the MIT License.


Created by GrzyboZAUR – feel free to contribute or fork the project!