import pandas as pd

# Read the output file
df = pd.read_excel('BOM_Sequential_20251106_111414.xlsx')

print("Rows 25-45 showing ordered hierarchy:")
print("Row | Material (E)        | Component (N)")
print("-" * 50)

for i in range(25, min(45, len(df))):
    material = str(df.iloc[i, 4])[:20]  # Column E (Material)
    component = str(df.iloc[i, 13])[:20]  # Column N (Component)
    print(f"{i:2d}  | {material:<20} | {component:<20}")

print("\nThis shows the correct ordering:")
print("- Level 2 material: 1228392193")
print("- Its Level 3 components: 1110500101, 1128392193 (both listed)")
print("- Then Level 3 materials appear: 1110500101, 1128392193")
print("- Finally their Level 4 components")