#!/usr/bin/env python3
"""
Test script untuk memverifikasi logika pencocokan nama
"""

import pandas as pd
import os

def test_name_matching():
    """Test pencocokan nama antara file wisuda dan list_pekerjaan"""
    
    # Load company lookup
    print("=== Testing Company Lookup ===")
    if os.path.exists('list_pekerjaan.xlsx'):
        df_pekerjaan = pd.read_excel('list_pekerjaan.xlsx')
        print(f"Loaded list_pekerjaan.xlsx with {len(df_pekerjaan)} rows")
        print(f"Columns: {list(df_pekerjaan.columns)}")
        
        # Create lookup dictionary with UPPERCASE keys
        lookup = {}
        for _, row in df_pekerjaan.iterrows():
            nama = str(row.get('Nama', '')).strip()
            nama_perusahaan = str(row.get('Nama Perusahaan', '')).strip()
            if nama and nama != 'nan' and nama_perusahaan and nama_perusahaan != 'nan':
                nama_upper = nama.upper()
                lookup[nama_upper] = nama_perusahaan
                print(f"Added: '{nama_upper}' -> '{nama_perusahaan}'")
        
        print(f"\nTotal lookup entries: {len(lookup)}")
        print(f"Sample entries: {dict(list(lookup.items())[:5])}")
    else:
        print("list_pekerjaan.xlsx not found!")
        return
    
    # Test with wisuda data
    print("\n=== Testing with Wisuda Data ===")
    
    # Test pagi data
    if os.path.exists('wisuda_pagi.xlsx'):
        df_pagi = pd.read_excel('wisuda_pagi.xlsx')
        print(f"Loaded wisuda_pagi.xlsx with {len(df_pagi)} students")
        
        if 'NAMA MAHASISWA' in df_pagi.columns:
            matches_found = 0
            total_students = 0
            
            for _, student in df_pagi.iterrows():
                nama = str(student.get('NAMA MAHASISWA', '')).strip()
                if nama and nama != 'nan':
                    total_students += 1
                    nama_upper = nama.upper()
                    perusahaan = lookup.get(nama_upper, None)
                    
                    if perusahaan:
                        matches_found += 1
                        print(f"MATCH: '{nama}' -> '{nama_upper}' -> '{perusahaan}'")
                    else:
                        print(f"NO MATCH: '{nama}' -> '{nama_upper}'")
            
            print(f"\nPagi Session Results:")
            print(f"Total students: {total_students}")
            print(f"Matches found: {matches_found}")
            print(f"Match rate: {(matches_found/total_students*100):.1f}%" if total_students > 0 else "No students")
    
    # Test siang data
    if os.path.exists('wisuda_siang.xlsx'):
        df_siang = pd.read_excel('wisuda_siang.xlsx')
        print(f"\nLoaded wisuda_siang.xlsx with {len(df_siang)} students")
        
        if 'NAMA MAHASISWA' in df_siang.columns:
            matches_found = 0
            total_students = 0
            
            for _, student in df_siang.iterrows():
                nama = str(student.get('NAMA MAHASISWA', '')).strip()
                if nama and nama != 'nan':
                    total_students += 1
                    nama_upper = nama.upper()
                    perusahaan = lookup.get(nama_upper, None)
                    
                    if perusahaan:
                        matches_found += 1
                        print(f"MATCH: '{nama}' -> '{nama_upper}' -> '{perusahaan}'")
                    else:
                        print(f"NO MATCH: '{nama}' -> '{nama_upper}'")
            
            print(f"\nSiang Session Results:")
            print(f"Total students: {total_students}")
            print(f"Matches found: {matches_found}")
            print(f"Match rate: {(matches_found/total_students*100):.1f}%" if total_students > 0 else "No students")

if __name__ == "__main__":
    test_name_matching()
