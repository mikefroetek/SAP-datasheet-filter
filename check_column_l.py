import pandas as pd
import openpyxl

# Read the BOM_Template_Aptiv file to see column L pattern
df = pd.read_excel('BOM_Template_Aptiv.xlsx')

print("Column L pattern in BOM_Template_Aptiv:")
print("Row | Column L (Item number)")
print("-" * 40)

for i in range(min(50, len(df))):
    col_l = str(df.iloc[i, 11])  # Column L is index 11
    col_e = str(df.iloc[i, 4])   # Column E (Material)
    col_n = str(df.iloc[i, 13])  # Column N (Component)
    print(f"{i:2d}  | {col_l:<20} | Mat: {col_e[:10]:<10} | Comp: {col_n[:10]:<10}")