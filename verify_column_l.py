import pandas as pd

# Read the new output file
df = pd.read_excel('BOM_Sequential_20251106_125305.xlsx')

print("Column L (Item number) verification:")
print("Row | Column L     | Material (E)     | Component (N)")
print("-" * 65)

for i in range(min(40, len(df))):
    col_l = str(df.iloc[i, 11])  # Column L (Item number)
    col_e = str(df.iloc[i, 4])   # Column E (Material) 
    col_n = str(df.iloc[i, 13])  # Column N (Component)
    print(f"{i:2d}  | {col_l:<12} | {col_e[:15]:<15} | {col_n[:15]:<15}")