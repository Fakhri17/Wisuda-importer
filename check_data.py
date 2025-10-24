import pandas as pd

# Check list_pekerjaan.xlsx
print("=== LIST PEKERJAAN ===")
try:
    df_pekerjaan = pd.read_excel('list_pekerjaan.xlsx')
    print(f"Rows: {len(df_pekerjaan)}")
    print(f"Columns: {list(df_pekerjaan.columns)}")
    print("First 5 rows:")
    print(df_pekerjaan.head())
    print("\nSample names:")
    if 'Nama' in df_pekerjaan.columns:
        print(df_pekerjaan['Nama'].head(10))
except Exception as e:
    print(f"Error: {e}")

print("\n=== WISUDA PAGI ===")
try:
    df_pagi = pd.read_excel('wisuda_pagi.xlsx')
    print(f"Rows: {len(df_pagi)}")
    print(f"Columns: {list(df_pagi.columns)}")
    print("Sample names:")
    if 'NAMA MAHASISWA' in df_pagi.columns:
        print(df_pagi['NAMA MAHASISWA'].head(10))
except Exception as e:
    print(f"Error: {e}")

print("\n=== WISUDA SIANG ===")
try:
    df_siang = pd.read_excel('wisuda_siang.xlsx')
    print(f"Rows: {len(df_siang)}")
    print(f"Columns: {list(df_siang.columns)}")
    print("Sample names:")
    if 'NAMA MAHASISWA' in df_siang.columns:
        print(df_siang['NAMA MAHASISWA'].head(10))
except Exception as e:
    print(f"Error: {e}")
