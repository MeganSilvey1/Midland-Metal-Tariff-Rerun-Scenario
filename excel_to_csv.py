import pandas as pd

# === CONFIGURATION ===
excel_file_path = "new/Bidsheet Master Consolidate Landed 12052025.xlsx"
sheet_name = "Sheet1"
csv_output_path = "new/Bidsheet Master Consolidate Landed 12052025.csv"

# === CONVERSION ===
print(f"Reading Excel file: {excel_file_path}")
df = pd.read_excel(excel_file_path, sheet_name=sheet_name, engine="openpyxl")
print(f"Saving as CSV: {csv_output_path}")
df.to_csv(csv_output_path, index=False)
print("Conversion complete.")

