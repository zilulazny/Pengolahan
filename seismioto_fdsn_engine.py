# %%
import glob
import numpy as np
import pygmt
import pandas as pd
import math
import matplotlib.pyplot as plt
import os
import tkinter as tk
import locale
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.shared import Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION, WD_ORIENT
from geopy.distance import geodesic
from openpyxl import load_workbook
from tkcalendar import DateEntry 
from tkinter import ttk, filedialog, messagebox
from config import *

# Fungsi untuk memilih file
def select_files():
    global selected_files
    selected_files = filedialog.askopenfilenames(title="Pilih File")
    if selected_files:
        messagebox.showinfo("File Terpilih", f"Jumlah file terpilih: {len(selected_files)}")

selected_files = []
tanggal_awal_gui = None
tanggal_akhir_gui = None

def set_input(files, tgl_awal, tgl_akhir):
    global selected_files
    global tanggal_awal_gui
    global tanggal_akhir_gui

    selected_files = files
    tanggal_awal_gui = tgl_awal
    tanggal_akhir_gui = tgl_akhir

def process_data():
    import xml.etree.ElementTree as ET
    global selected_files, tanggal_awal, tanggal_akhir, nama_folder
    CITY_FILE = CITY

    tanggal_awal = pd.to_datetime(
        tanggal_awal_gui,
        format='%d/%m/%Y'
    )

    tanggal_akhir = pd.to_datetime(
        tanggal_akhir_gui,
        format='%d/%m/%Y'
    )

    if not selected_files:
        messagebox.showerror("Error", "Pilih file XML terlebih dahulu!")
        return None, None

    ns = {'sc': 'http://geofon.gfz-potsdam.de/ns/seiscomp3-schema/0.13'}
    combined_data = pd.DataFrame()

    def load_city_list():
        L = []
        if not os.path.exists(CITY_FILE):
            print("[WARN] kabkota.txt tidak ditemukan")
            return L
        with open(CITY_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                p = line.split(";")
                if len(p) < 3:
                    continue
                try:
                    lat = float(p[0])
                    lon = float(p[1])
                except Exception:
                    continue
                name = ";".join(p[2:])
                L.append((lat, lon, name))
        return L

    CITY_LIST = load_city_list()

    def calculate_azimuth(lat1, lon1, lat2, lon2):
        rlat1 = math.radians(lat1)
        rlon1 = math.radians(lon1)
        rlat2 = math.radians(lat2)
        rlon2 = math.radians(lon2)
        dlon = rlon2 - rlon1
        x = math.sin(dlon) * math.cos(rlat2)
        y = (math.cos(rlat1) * math.sin(rlat2) -
            math.sin(rlat1) * math.cos(rlat2) * math.cos(dlon))
        az = math.degrees(math.atan2(x, y))
        if az < 0:
            az += 360
        return az

    def azimuth_to_direction(a: float) -> str:
        if 0 <= a < 22.5 or 337.5 <= a < 360:
            return "Utara"
        if 22.5 <= a < 67.5:
            return "Timur Laut"
        if 67.5 <= a < 112.5:
            return "Timur"
        if 112.5 <= a < 157.5:
            return "Tenggara"
        if 157.5 <= a < 202.5:
            return "Selatan"
        if 202.5 <= a < 247.5:
            return "Barat Daya"
        if 247.5 <= a < 292.5:
            return "Barat"
        if 292.5 <= a < 337.5:
            return "Barat Laut"
        return "-"

    def nearest_city(lat, lon) -> str:
        if not CITY_LIST:
            return "-"
        try:
            lat = float(lat)
            lon = float(lon)
        except Exception:
            return "-"

        best_d = float("inf")
        best_city = "-"
        best_dir = "-"

        for clat, clon, cname in CITY_LIST:
            try:
                d = geodesic((lat, lon), (clat, clon)).km
            except Exception:
                continue
            if d < best_d:
                best_d = d
                best_city = cname
                az = calculate_azimuth(clat, clon, lat, lon)
                best_dir = azimuth_to_direction(az)

        if best_city == "-":
            return "-"

        jarak = int(round(best_d))
        arah = best_dir.replace(" ", "")  # TimurLaut, BaratDaya, dll
        return f"{jarak} km {arah} {best_city}"
    
        # ===== FUNGSI: DEDUPLIKASI GEMPA =====
    def deduplicate_earthquakes(df, time_tolerance_seconds=5):
        """
        Menghapus gempa duplikat berdasarkan tanggal + waktu (±5 detik)
        Menyimpan data pertama (index terkecil) jika ada duplikat
        """
        if df.empty:
            return df

        df = df.sort_values(by=['Tanggal', 'Waktu (WIB)']).reset_index(drop=True)
        df['Datetime'] = pd.to_datetime(
            df['Tanggal'].astype(str) + ' ' + df['Waktu (WIB)'],
            errors='coerce'
        )

        duplicates_to_drop = []

        for i in range(len(df)):
            if i in duplicates_to_drop:
                continue

            current_time = df.loc[i, 'Datetime']
            current_date = df.loc[i, 'Tanggal']

            # Bandingkan dengan gempa berikutnya
            for j in range(i + 1, len(df)):
                if j in duplicates_to_drop:
                    continue

                compare_time = df.loc[j, 'Datetime']
                compare_date = df.loc[j, 'Tanggal']

                # Jika tanggal berbeda, skip
                if current_date != compare_date:
                    break

                # Cek perbedaan waktu
                time_diff = abs((current_time - compare_time).total_seconds())

                if time_diff <= time_tolerance_seconds:
                    # Tandai untuk dihapus (keep yang pertama)
                    duplicates_to_drop.append(j)

        # Hapus duplikat
        df_deduplicated = df.drop(duplicates_to_drop).reset_index(drop=True)
        df_deduplicated = df_deduplicated.drop(columns=['Datetime'])

        return df_deduplicated

    # ===== FUNGSI: MERGE DENGAN DIRASAKAN =====
    def merge_with_dirasakan(combined_data, dirasakan_file, time_tolerance_seconds=5):
        """
        Merge data gempa dengan dirasakan.xlsx

        Aturan:
        1. Match berdasarkan tanggal + waktu (±5 detik)
        2. Jika match -> update kolom Dirasakan
        3. Jika tidak match -> tambahkan data dari Excel ke combined_data
        """

        import os
        import pandas as pd
        from tkinter import messagebox

        if not os.path.exists(dirasakan_file):
            messagebox.showerror(
                "Error",
                f"File '{dirasakan_file}' tidak ditemukan!"
            )
            return combined_data

        # ==========================
        # BACA FILE EXCEL
        # ==========================
        existing_data = pd.read_excel(dirasakan_file)

        # Pastikan kolom wajib ada
        required_columns = ['Tanggal', 'Waktu (WIB)']
        for col in required_columns:
            if col not in existing_data.columns:
                messagebox.showerror(
                    "Error",
                    f"Kolom '{col}' tidak ditemukan pada file dirasakan.xlsx"
                )
                return combined_data

        # ==========================
        # NORMALISASI TANGGAL & WAKTU
        # ==========================
        existing_data['Tanggal'] = pd.to_datetime(
            existing_data['Tanggal'],
            errors='coerce'
        )

        existing_data['Waktu (WIB)'] = (
            existing_data['Waktu (WIB)']
            .astype(str)
            .str.strip()
        )

        existing_data['Datetime'] = pd.to_datetime(
            existing_data['Tanggal'].astype(str)
            + ' '
            + existing_data['Waktu (WIB)'],
            errors='coerce'
        )

        combined_data['Tanggal'] = pd.to_datetime(
            combined_data['Tanggal'],
            errors='coerce'
        )

        combined_data['Waktu (WIB)'] = (
            combined_data['Waktu (WIB)']
            .astype(str)
            .str.strip()
        )

        combined_data['Datetime'] = pd.to_datetime(
            combined_data['Tanggal'].astype(str)
            + ' '
            + combined_data['Waktu (WIB)'],
            errors='coerce'
        )

        # ==========================
        # FLAG SEMENTARA
        # ==========================
        combined_data['from_excel'] = False

        indices_to_replace = []
        matched_existing = set()

        # ==========================
        # CARI DATA YANG MATCH
        # ==========================
        for idx_existing, row_existing in existing_data.iterrows():

            if pd.isna(row_existing['Datetime']):
                continue

            for idx_combined, row_combined in combined_data.iterrows():

                if pd.isna(row_combined['Datetime']):
                    continue

                if row_combined['from_excel']:
                    continue

                time_diff = abs(
                    (
                        row_combined['Datetime']
                        - row_existing['Datetime']
                    ).total_seconds()
                )

                if time_diff <= time_tolerance_seconds:

                    indices_to_replace.append(
                        (idx_combined, idx_existing)
                    )

                    matched_existing.add(idx_existing)

                    combined_data.loc[
                        idx_combined,
                        'from_excel'
                    ] = True

                    break

        # ==========================
        # UPDATE KOLOM DIRASAKAN
        # ==========================
        for idx_combined, idx_existing in indices_to_replace:

            if 'Dirasakan' in existing_data.columns:

                combined_data.loc[
                    idx_combined,
                    'Dirasakan'
                ] = existing_data.loc[
                    idx_existing,
                    'Dirasakan'
                ]

        # ==========================
        # TAMBAHKAN GEMPA YANG
        # TIDAK ADA DI COMBINED_DATA
        # ==========================
        rows_to_append = []

        for idx_existing, row_existing in existing_data.iterrows():

            if idx_existing in matched_existing:
                continue

            new_row = {}

            for col in combined_data.columns:

                if col in row_existing.index:
                    new_row[col] = row_existing[col]
                else:
                    new_row[col] = ''

            rows_to_append.append(new_row)

        if rows_to_append:

            combined_data = pd.concat(
                [
                    combined_data,
                    pd.DataFrame(rows_to_append)
                ],
                ignore_index=True
            )

        # ==========================
        # HAPUS KOLOM BANTU
        # ==========================
        columns_to_drop = []

        if 'from_excel' in combined_data.columns:
            columns_to_drop.append('from_excel')

        if 'Datetime' in combined_data.columns:
            columns_to_drop.append('Datetime')

        combined_data = combined_data.drop(
            columns=columns_to_drop,
            errors='ignore'
        )

        # ==========================
        # SORT BERDASARKAN WAKTU
        # ==========================
        combined_data['Datetime'] = pd.to_datetime(
            combined_data['Tanggal'].astype(str)
            + ' '
            + combined_data['Waktu (WIB)'].astype(str),
            errors='coerce'
        )

        combined_data = combined_data.sort_values(
            by='Datetime'
        ).reset_index(drop=True)

        combined_data = combined_data.drop(
            columns=['Datetime'],
            errors='ignore'
        )

        return combined_data

    # Loop tiap file XML
    for file_path in selected_files:
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            data_csv = []

            for event in root.findall(".//sc:event", ns):
                event_id = event.attrib.get("publicID")
                preferred_origin_id = event.findtext("sc:preferredOriginID", namespaces=ns)
                preferred_magnitude_id = event.findtext("sc:preferredMagnitudeID", namespaces=ns)
                deskripsi = event.find("sc:description/sc:text", ns)
                deskripsi_lokasi = deskripsi.text if deskripsi is not None else ""

                origin = root.find(f".//sc:origin[@publicID='{preferred_origin_id}']", ns)
                if origin is None:
                    continue

                # ===== FILTER: Hanya ambil origin dengan evaluationMode = manual =====
                eval_mode = origin.findtext("sc:evaluationMode", namespaces=ns)
                if eval_mode is None or eval_mode.lower() != "manual":
                    # Skip origin jika evaluationMode bukan manual
                    continue

                # Ambil data utama
                waktu_utc = origin.findtext("sc:time/sc:value", namespaces=ns)
                latitude = float(origin.findtext("sc:latitude/sc:value", namespaces=ns))
                longitude = float(origin.findtext("sc:longitude/sc:value", namespaces=ns))
                depth = float(origin.findtext("sc:depth/sc:value", namespaces=ns))

                # Konversi waktu ke UTC+7
                t_local = datetime.strptime(waktu_utc, "%Y-%m-%dT%H:%M:%S.%fZ") + timedelta(hours=7)
                tanggal = t_local.strftime("%Y-%m-%d")
                jam = t_local.strftime("%H:%M:%S")

                # Parse magnitude dengan error handling
                magnitude = None
                if preferred_magnitude_id is not None:
                    try:
                        mag_elem = root.find(f".//sc:magnitude[@publicID='{preferred_magnitude_id}']/sc:magnitude/sc:value", ns)
                        if mag_elem is not None:
                            magnitude = round(float(mag_elem.text), 1)
                    except (ValueError, AttributeError):
                        magnitude = None

                if magnitude is None:
                    continue

                # Simpan hasil untuk CSV
                data_csv.append({
                    "Tanggal": tanggal,
                    "Waktu (WIB)": jam,
                    "Lintang": round(latitude, 2),
                    "Bujur": round(longitude, 2),
                    "Kedalaman": int(depth),
                    "Magnitude": magnitude,
                    "Keterangan": deskripsi_lokasi,
                    "Dirasakan": ''
                })

            # Buat DataFrame dari CSV
            parsed_data = pd.DataFrame(data_csv)
            parsed_data['Tanggal'] = pd.to_datetime(parsed_data['Tanggal'], errors='coerce')
            parsed_data["Lintang"] = parsed_data["Lintang"].round(2)
            parsed_data["Bujur"] = parsed_data["Bujur"].round(2)
            parsed_data["Kedalaman"] = parsed_data["Kedalaman"].round(0).astype("Int64")  # int tapi bisa NaN
            parsed_data["Magnitude"] = parsed_data["Magnitude"].round(1)

            # Filter tanggal & koordinat & magnitude
            mask = (
                (parsed_data['Tanggal'] >= tanggal_awal) &
                (parsed_data['Tanggal'] <= tanggal_akhir) &
                (parsed_data['Bujur'].astype(float) >= 114) &    # 113.95 / 114
                (parsed_data['Bujur'].astype(float) <= 116.3) &    # 117.31 / 116.3
                (parsed_data['Lintang'].astype(float) >= -11.5) &  # -13.65 / -11.5
                (parsed_data['Lintang'].astype(float) <= -7.05)     # -6.68 / -7.05
            )
            
            filtered_data = parsed_data[mask].copy()
            # ===== INI YANG BARU - GANTI KETERANGAN =====
            print("Menghitung jarak kota terdekat...")
            filtered_data["Keterangan"] = filtered_data.apply(
                lambda row: nearest_city(row["Lintang"], row["Bujur"]), 
                axis=1
            )
            print("Selesai!")
            combined_data = pd.concat([combined_data, filtered_data], ignore_index=True)

        except Exception as e:
            messagebox.showerror("Error", f"Error parsing {file_path}: {str(e)}")
            continue

    if combined_data.empty:
        messagebox.showwarning("Warning", "Tidak ada data yang sesuai filter!")
        return

    # ===== PROSES UTAMA =====

    # 1. Deduplikasi gempa
    combined_data = deduplicate_earthquakes(combined_data, time_tolerance_seconds=5)

    # 2. Merge dengan dirasakan.xlsx
    existing_file = DIRASAKAN_FILE
    combined_data = merge_with_dirasakan(combined_data, existing_file, time_tolerance_seconds=5)

    # 3. Sort & format output
    combined_data['Tanggal'] = pd.to_datetime(combined_data['Tanggal'])
    combined_data['Datetime'] = pd.to_datetime(
        combined_data['Tanggal'].astype(str) + ' ' + combined_data['Waktu (WIB)'],
        errors='coerce'
    )
    combined_data = combined_data.sort_values(by='Datetime').reset_index(drop=True)
    combined_data = combined_data.drop(columns=['Datetime'])
    combined_data['Tanggal'] = combined_data['Tanggal'].dt.date

    # 4. Reorder columns & add No
    column_order = ['No', 'Tanggal', 'Waktu (WIB)', 'Lintang', 'Bujur', 'Kedalaman', 'Magnitude', 'Keterangan', 'Dirasakan']
    combined_data.insert(0, 'No', range(1, len(combined_data) + 1))
    combined_data = combined_data[column_order]

    # 5. Save output
    combined_data.to_csv(DATALENGKAP_FILE, index=False)
    
    # HAPUS yang lama, GANTI dengan ini:
    bulan_id = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
        5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
        9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }

    if tanggal_awal.month == tanggal_akhir.month:
        nama_folder = f"{tanggal_awal.day:02d} - {tanggal_akhir.day:02d} {bulan_id[tanggal_awal.month]} {tanggal_awal.year}"
    else:
        nama_folder = f"{tanggal_awal.day:02d} {bulan_id[tanggal_awal.month]} - {tanggal_akhir.day:02d} {bulan_id[tanggal_akhir.month]} {tanggal_akhir.year}"

    folder_output = os.path.join(OUTPUT_DIR, nama_folder)

    os.makedirs(folder_output, exist_ok=True)

    output_file = os.path.join(folder_output, f'Data Gempabumi Periode {nama_folder}.xlsx')
    combined_data.to_excel(output_file, index=False)

    # ===== TAMBAH SESUDAH PENYIMPANAN EXCEL, SEBELUM MESSAGEBOX =====
    data_peta_folder = os.path.join(folder_output, "data peta")
    os.makedirs(data_peta_folder, exist_ok=True)

    # Klasifikasi berdasarkan kedalaman
    dangkal = combined_data[(combined_data['Kedalaman'] >= 0) & (combined_data['Kedalaman'] <= 60)].copy()
    menengah = combined_data[(combined_data['Kedalaman'] > 60) & (combined_data['Kedalaman'] < 300)].copy()
    dalam = combined_data[(combined_data['Kedalaman']>= 300)]
    # dalam_header = pd.DataFrame(columns=combined_data.columns)
    # dalam_header.to_csv(os.path.join(data_peta_folder, 'dalam.csv'), index=False)

    # Simpan file CSV
    if not dangkal.empty:
        dangkal.to_csv(os.path.join(data_peta_folder, 'dangkal.csv'), index=False)
    if not menengah.empty:
        menengah.to_csv(os.path.join(data_peta_folder, 'menengah.csv'), index=False)
    if not dalam.empty:
        dalam.to_csv(os.path.join(data_peta_folder, 'dalam.csv'), index=False)

    return nama_folder

def print_analysis():
    global nama_folder
    def read_csv_and_analyze(tanggal_awal, tanggal_akhir):
        
        # Step 1: Read CSV file
        file_path = DATALENGKAP_FILE
        data = pd.read_csv(file_path)
        
        # Konversi kolom Tanggal ke datetime dan ambil hanya tanggalnya
        data['Tanggal'] = pd.to_datetime(data['Tanggal']).dt.date
        tanggal_awal = pd.to_datetime(tanggal_awal).date()
        tanggal_akhir = pd.to_datetime(tanggal_akhir).date()
          
        # Filter data berdasarkan tanggal yang dipilih
        data = data[(data['Tanggal'] >= tanggal_awal) & (data['Tanggal'] <= tanggal_akhir)]

        # Drop the 'Magnitude Category' and 'Depth Category' columns
        output_data = data.drop(columns=['Magnitude Category', 'Depth Category'], errors='ignore')

        # Create magnitude and depth categories
        categories_magnitude = ['M≤3', '3<M<5', 'M≥5']

        data['Magnitude'] = pd.to_numeric(
            data['Magnitude'],
            errors='coerce'
        )

        conditions = [
            (data['Magnitude'] > 0) & (data['Magnitude'] <= 3),
            (data['Magnitude'] > 3) & (data['Magnitude'] < 5),
            (data['Magnitude'] >= 5)
        ]

        data['Magnitude Category'] = np.select(
            conditions,
            categories_magnitude,
            default=None
        )
        #data['Magnitude Category'] = pd.cut(data['Magnitude'], bins=[0, 3, 5, float('inf')], labels=categories_magnitude, right=True)

        categories_depth = ['D≤60 km', '60<D<300 km', 'D≥300 km']
        data['Kedalaman'] = pd.to_numeric(data['Kedalaman'], errors='coerce')
        data['Depth Category'] = pd.cut(data['Kedalaman'], bins=[0, 60, 300, float('inf')], labels=categories_depth, right=True)       
        
        # Group data by day and categories for analysis
        # magnitude_summary = (data.groupby([data['Tanggal'], 'Magnitude Category']).size().unstack(fill_value=0).reindex(index=all_dates, fill_value=0).reindex(columns=categories_magnitude, fill_value=0))
        # depth_summary = (data.groupby([data['Tanggal'], 'Depth Category']).size().unstack(fill_value=0).reindex(index=all_dates, fill_value=0).reindex(columns=categories_depth, fill_value=0))
        
        magnitude_summary = (data.groupby([data['Tanggal'], 'Magnitude Category'])
                            .size()
                            .unstack(fill_value=0)
                            .reindex(index=(pd.date_range(start=tanggal_awal, end=tanggal_akhir, freq='D')), fill_value=0)
                            .reindex(columns=categories_magnitude, fill_value=0))

        depth_summary = (data.groupby([data['Tanggal'], 'Depth Category'])
                        .size()
                        .unstack(fill_value=0)
                        .reindex(index=(pd.date_range(start=tanggal_awal, end=tanggal_akhir, freq='D')), fill_value=0)
                        .reindex(columns=categories_depth, fill_value=0))

        # Add totals row and columns for 'Gempa Dirasakan' and 'Gempa Merusak'
        magnitude_summary['Jumlah Total'] = magnitude_summary.sum(axis=1)
        magnitude_summary['Gempa Dirasakan'] = 0
        magnitude_summary['Gempa Merusak'] = 0
        magnitude_summary.loc['Jumlah Gempa'] = magnitude_summary.sum()

        depth_summary['Jumlah Total'] = depth_summary.sum(axis=1)
        depth_summary['Gempa Dirasakan'] = 0
        depth_summary['Gempa Merusak'] = 0
        depth_summary.loc['Jumlah Gempa'] = depth_summary.sum()

        # Mengisi kolom 'Gempa Dirasakan' jika kolom 'Dirasakan' terisi
        for index, row in data.iterrows():
            if pd.notna(row['Dirasakan']) and row['Dirasakan'].strip() != '':
                if row['Tanggal'] in magnitude_summary.index:
                    magnitude_summary.at[row['Tanggal'], 'Gempa Dirasakan'] += 1
                if row['Tanggal'] in depth_summary.index:
                    depth_summary.at[row['Tanggal'], 'Gempa Dirasakan'] += 1

        magnitude_summary.at['Jumlah Gempa', 'Gempa Dirasakan'] = magnitude_summary['Gempa Dirasakan'].sum()
        depth_summary.at['Jumlah Gempa', 'Gempa Dirasakan'] = depth_summary['Gempa Dirasakan'].sum()

        # Untuk menangani row 'Jumlah Gempa' yang bukan tanggal
        magnitude_summary.index = [x.strftime('%Y-%m-%d') if isinstance(x, pd.Timestamp) else x for x in magnitude_summary.index]
        depth_summary.index = [x.strftime('%Y-%m-%d') if isinstance(x, pd.Timestamp) else x for x in depth_summary.index]

        # Save the analysis to Excel
        folder_output = os.path.join(OUTPUT_DIR, nama_folder)
        output_file = os.path.join(folder_output, f'Analisis Gempabumi Periode {nama_folder}.xlsx')
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            output_data.to_excel(writer, sheet_name='Earthquake Data', index=False)
            magnitude_summary.to_excel(writer, sheet_name='Magnitude Summary')
            depth_summary.to_excel(writer, sheet_name='Depth Summary')
            plot_charts(magnitude_summary, depth_summary, writer, nama_folder, tanggal_awal, tanggal_akhir)

        # Autofit columns
        autofit_columns(output_file, 'Earthquake Data')
        autofit_columns(output_file, 'Magnitude Summary')
        autofit_columns(output_file, 'Depth Summary')

        #messagebox.showinfo("Info", f"Analisis berhasil disimpan di:\n{output_file}")

    def plot_charts(magnitude_summary, depth_summary, writer, nama_folder, tanggal_awal, tanggal_akhir,):
        folder_output = os.path.join(OUTPUT_DIR, nama_folder)
        # ================== MAGNITUDE BAR CHART ==================
        categories = ['M≤3', '3<M<5', 'M≥5']
        colors = ['green', 'yellow', 'red']
        
        # Remove total row
        magnitude_summary_without_total = magnitude_summary.drop(index='Jumlah Gempa', errors='ignore')
        
        # Reindex with all dates and fill missing with 0
        magnitude_summary_without_total.index = pd.to_datetime(magnitude_summary_without_total.index)
        magnitude_summary_without_total.index = magnitude_summary_without_total.index.strftime('%d-%m-%Y')

        # Plot setup
        fig, ax = plt.subplots(figsize=(10, 6))
        bar_width = 0.25
        x = np.arange(len(magnitude_summary_without_total.index))

        # Plot each category
        for i, category in enumerate(categories):
            ax.bar(x + i * bar_width, 
                magnitude_summary_without_total[category], 
                bar_width, 
                label=category, 
                color=colors[i], 
                edgecolor='black')

        # Chart formatting
        ax.set_title('Frekuensi Gempa Berdasarkan Magnitudo',fontsize=14, pad=20)
        ax.set_xlabel('Tanggal', fontsize=12)
        ax.set_ylabel('Jumlah Kejadian', fontsize=12)
        
        # Set x-ticks to show all dates
        ax.set_xticks(x + bar_width)
        
        # Update x-tick labels dengan format baru
        date_labels = [idx.strftime('%d %b') if hasattr(idx, 'strftime') else str(idx)[:5] 
               for idx in magnitude_summary_without_total.index]
        ax.set_xticklabels(date_labels, rotation=30, ha='right', fontsize=8)    
        
        ax.legend(title='Kategori Magnitudo', fontsize=10)
        ax.grid(axis='y', linestyle='--', alpha=0.7)

        # Add value labels
        for p in ax.patches:
            height = p.get_height()
            if height > 0:  # Hanya tampilkan label jika nilai > 0
                ax.annotate(f'{int(height)}', 
                        (p.get_x() + p.get_width() / 2, height),
                        ha='center', va='bottom', fontsize=8)

        plt.tight_layout()
        plt.savefig(os.path.join(folder_output, 'diagbat_mag.png'), dpi=300)
        plt.close()

        # Insert to Excel
        worksheet_mag = writer.sheets['Magnitude Summary']
        worksheet_mag.insert_image('I2', os.path.join(folder_output, 'diagbat_mag.png'), 
                                {'x_scale': 0.5, 'y_scale': 0.5})

        # Buat diagram lingkaran magnitude
        magnitude_total = magnitude_summary.loc['Jumlah Gempa', ['M≤3', '3<M<5', 'M≥5']]
        labels = ['M≤3', '3<M<5', 'M≥5']
        colors = ['green', 'yellow', 'red']

        # Membuat figure dan pie chart
        fig, ax = plt.subplots(figsize=(6, 6))

        # Fungsi untuk menampilkan persentase dan jumlah kejadian di setiap segmen
        def func(pct, allvals):
            absolute = int(np.round(pct/100.*np.sum(allvals)))
            if absolute > 0:
                return f'{pct:.1f}%\n({absolute} kejadian)'
            else:
                return ''  # Jika tidak ada kejadian, label tidak ditampilkan

        # Membuat pie chart
        wedges, texts, autotexts = ax.pie(
            magnitude_total,
            labels=labels,
            colors=colors,
            startangle=85,
            autopct=lambda pct: func(pct, magnitude_total),  # Menampilkan persentase dan jumlah kejadian
            pctdistance=0.875,  # Jarak persentase dari pusat pie
            labeldistance=9  # Jarak label dari pusat pie (aslinya 1.17)
        )

        # Mengatur ukuran font dan latar belakang putih untuk label dan persentase
        for text in texts:
            text.set_fontsize(12)
            text.set_bbox(dict(facecolor='white', edgecolor='none', boxstyle='round,pad=0.2'))

        for autotext in autotexts:
            autotext.set_fontsize(12)
            autotext.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.2'))

        plt.title('Distribusi Magnitudo', fontsize=11)
        ax.legend(title='Kategori Magnitudo', fontsize=7, title_fontsize='9', loc='lower right')
        plt.tight_layout()

        plt.savefig(os.path.join(folder_output, 'diaglingk_mag.png'))

        # Insert the magnitude pie chart into Excel
        worksheet_mag.insert_image('P2', os.path.join(folder_output, 'diaglingk_mag.png'), {'x_scale': 0.5, 'y_scale': 0.5})
        
        # ================== DEPTH BAR CHART ==================
        categories_depth = ['D≤60 km', '60<D<300 km', 'D≥300 km']
        colors_depth = ['red', 'yellow', 'green']
        
       # Remove total row
        depth_summary_without_total = depth_summary.drop(index='Jumlah Gempa', errors='ignore')
        
        # Reindex with all dates and fill missing with 0
        depth_summary_without_total.index = pd.to_datetime(depth_summary_without_total.index)
        depth_summary_without_total.index = depth_summary_without_total.index.strftime('%d-%m-%Y')

        # Plot setup
        fig, ax = plt.subplots(figsize=(10, 6))
        bar_width = 0.25
        x = np.arange(len(depth_summary_without_total.index))
        
        # Plot each category
        for i, category in enumerate(categories_depth):
            ax.bar(x + i * bar_width, 
                depth_summary_without_total[category], 
                bar_width, 
                label=category, 
                color=colors_depth[i], 
                edgecolor='black')

        # Chart formatting
        ax.set_title('Frekuensi Gempa Berdasarkan Kedalaman' ,fontsize=14, pad=20)
        ax.set_xlabel('Tanggal', fontsize=12)
        ax.set_ylabel('Jumlah Kejadian', fontsize=12)
        
        # Set x-ticks to show all dates
        ax.set_xticks(x + bar_width)
        
        # Update x-tick labels dengan format baru
        date_labels = [idx.strftime('%d %b') if hasattr(idx, 'strftime') else str(idx)[:5] 
               for idx in depth_summary_without_total.index]
        ax.set_xticklabels(date_labels, rotation=30, ha='right', fontsize=8)  
        
        ax.legend(title='Kategori Kedalaman', fontsize=10)
        ax.grid(axis='y', linestyle='--', alpha=0.7)

        for p in ax.patches:
            height = p.get_height()
            if height > 0:  # Hanya tampilkan label jika nilai > 0
                ax.annotate(f'{int(height)}', 
                        (p.get_x() + p.get_width() / 2, height),
                        ha='center', va='bottom', fontsize=8)

        plt.tight_layout()
        plt.savefig(os.path.join(folder_output, 'diagbat_depth.png'), dpi=300)
        plt.close()

        worksheet_depth = writer.sheets['Depth Summary']
        worksheet_depth.insert_image('I2', os.path.join(folder_output, 'diagbat_depth.png'), 
                                {'x_scale': 0.5, 'y_scale': 0.5})

        # buat digram lingkaran depth
        depth_total = depth_summary.loc['Jumlah Gempa', ['D≤60 km', '60<D<300 km', 'D≥300 km']]
        labels = ['D≤60 km', '60<D<300 km', 'D≥300 km']
        colors = ['red', 'yellow', 'green']
        
        # Membuat figure dan pie chart
        fig, ax = plt.subplots(figsize=(6, 6))
        
        # Fungsi untuk menampilkan persentase dan jumlah kejadian di setiap segmen
        def func(pct, allvals):
            absolute = int(np.round(pct/100.*np.sum(allvals)))
            if absolute > 0:
                return f'{pct:.1f}%\n({absolute} kejadian)'
            else:
                return ''  # Jika tidak ada kejadian, label tidak ditampilkan
                  
        # Membuat pie chart
        wedges, texts, autotexts = ax.pie(
            depth_total,
            labels=labels,
            colors=colors,
            startangle=85,
            autopct=lambda pct: func(pct, magnitude_total),  # Menampilkan persentase dan jumlah kejadian
            pctdistance=0.975,  # Jarak persentase dari pusat pie
            labeldistance=9  # Jarak label dari pusat pie (aslinya 1.15)
        )

        # Mengatur ukuran font dan latar belakang putih untuk label dan persentase
        for text in texts:
            text.set_fontsize(12)
            text.set_bbox(dict(facecolor='white', edgecolor='none', boxstyle='round,pad=0.2'))

        for autotext in autotexts:
            autotext.set_fontsize(12)
            autotext.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.2'))

        plt.title('Distribusi Kedalaman', fontsize=11)
        ax.legend(title='Kategori Kedalaman', fontsize=7, title_fontsize='9', loc='lower right')
        plt.tight_layout()
        plt.savefig(os.path.join(folder_output, 'diaglingk_depth.png'))

        # Insert the magnitude pie chart into Excel
        worksheet_depth.insert_image('P2', os.path.join(folder_output, 'diaglingk_depth.png'), {'x_scale': 0.5, 'y_scale': 0.5})   
    
    def autofit_columns(file_path, sheet_name):
        """
        Menyesuaikan lebar kolom di Excel berdasarkan panjang konten.
        """
        workbook = load_workbook(file_path)  # Membuka workbook yang telah disimpan
        worksheet = workbook[sheet_name]  # Mendapatkan worksheet berdasarkan nama sheet

        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter  # Mendapatkan huruf kolom, misalnya 'A', 'B', dst.
            for cell in col:
                try:
                    # Mencari panjang maksimum dari setiap kolom
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)  # Menambah padding
            worksheet.column_dimensions[column].width = adjusted_width  # Mengatur lebar kolom

        workbook.save(file_path)  # Menyimpan workbook setelah penyesuaian
        workbook.close()  # Menutup workbook untuk memastikan semua perubahan disimpan
    
    # Panggil fungsi utama dengan tanggal yang dipilih
    read_csv_and_analyze(tanggal_awal, tanggal_akhir)

# def print_map():
#    #membuat file .dat
#     input=DATALENGKAP_FILE
#     df = pd.read_csv(input)
#     global nama_folder

#     # Memilih kolom yang diinginkan
#     selected_columns = df[['Bujur','Lintang','Kedalaman', 'Magnitude']]

#     # Menulis data ke file DAT
#     with open(INPUT_PETA, 'w') as file:
#         for index, row in selected_columns.iterrows():
#             line = '  '.join(row.astype(str))  # Menggabungkan semua elemen baris menjadi satu string dengan spasi sebagai pemisah
#             file.write(line + '\n')
            
#     # Membaca data dari hasil2.csv
#     data = pd.read_csv(input)
#     pygmt.config(PS_PAGE_COLOR="white")

#     # Buat figure
#     fig = pygmt.Figure()

#     # Buat color palette untuk kedalaman gempa
#     pygmt.makecpt(cmap="geo", series=[-7000, 4000])
#     pygmt.makecpt(cmap="red,yellow,green", series="0,60,300,1000", output="quakes.cpt")
#     # start_date_formatted = start_date_entry.get().replace('/', '-')  # Ganti format tanggal
#     # end_date_formatted = end_date_entry.get().replace('/', '-')
#     #title_text = f"Peta Seismisitas Wilayah Bali dan Sekitarnya \nPeriode {start_date_formatted} hingga {end_date_formatted}"
#     data['Magnitude'] = pd.to_numeric(data['Magnitude'], errors='coerce')  # Ganti nilai tidak valid dengan NaN

#     # Menambahkan peta dasar dengan grid dan garis pantai
#     fig.grdimage(
#         grid=EARTH_RELIEF,
#         region=[113.9, 117.31, -12, -7],
#         projection="M17c",
#         # shading=True,
#         frame=["xa2g2", "ya2g2"])#'+t"Peta Seismisitas"'])

#     fig.plot(data=str(SESAR_1), pen="1.5,black",fill='black')
#     fig.plot(data=str(SESAR_2), pen="1.5,black")
    
#     # Plot gempa yang tidak dirasakan (kolom 'Dirasakan' kosong)
#     gempa_tidak_dirasakan = data[data['Dirasakan'].isna()]
#     if not gempa_tidak_dirasakan.empty:
#         fig.plot(
#             x=gempa_tidak_dirasakan['Bujur'],
#             y=gempa_tidak_dirasakan['Lintang'],
#             size=0.08*gempa_tidak_dirasakan['Magnitude'],
#             style="cc",  # Lingkaran
#             fill=gempa_tidak_dirasakan['Kedalaman'].astype(float),  # Warna berdasarkan kedalaman
#             cmap="quakes.cpt",  # Gunakan colormap yang telah ditentukan
#             pen="black"
#         )
#     else:
#         print("ada gempa bumi yang dirasakan untuk diplot.")

#     # Plot gempa yang dirasakan (kolom 'Dirasakan' terisi)
#     gempa_dirasakan = data[~data['Dirasakan'].isna()]
#     if not gempa_dirasakan.empty:
#         fig.plot(
#             x=gempa_dirasakan['Bujur'],
#             y=gempa_dirasakan['Lintang'],
#             size=0.1*gempa_dirasakan['Magnitude'],
#             style="a0.5c",  # Simbol bintang
#             fill="orange",  # Warna bintang merah
#             pen="1p,black"  # Garis tepi hitam dengan ketebalan 1 point
#         )
#     else:
#         print("tidak ada gempa bumi yang dirasakan untuk diplot.")
    
#     fig.plot(x=[114.8, 115.4], y=[-11.5, -7.5], projection="M", pen=2)
#     fig.text(x=114.9, y=-11.6, text="A", font="18,Helvetica")
#     fig.text(x=115.55, y=-7.5, text="A1", font="18,Helvetica")
  
#     # Menambahkan colorbar
#     fig.image(str(LOGO_BMKG), position='n1.1/0.85+w1.1i')
#     fig.image(str(MTANGAN), position='n1.08/0.73+w1.1i')
#     fig.basemap(map_scale="n1.16/0.7+w100k+f0.1p+lKm")
#     fig.colorbar(
#         frame=["x+lKetinggian", "y"],  # Label untuk sumbu x dan y
#         position="n1.03/0.51+w4c/0.5c")  # Posisi colorbar (JMR = tengah-kanan), lebar 0.5 cm, panjang 10 cm
#     with fig.inset(position="n1.03/0.18+w5c/2.5c", box="+pblack"):
#         fig.coast(
#             region=[97, 140, -15, 7],
#             shorelines='0.5p,black',
#             projection="M5c",
#             land="green",
#             water="lightblue"
#         )
#         fig.plot(x=[117.31, 113.21, 113.21, 117.31, 117.31], y=[-7, -7, -12, -12, -7], pen="1p,red")
#     fig.legend(spec=str(LEGENDA), position="n1.03/0.0+w5c",box="+gwhite+p1p,black")

#         # Membuat cross section
#     pygmt.project(
#         data=INPUT_PETA,
#         unit=True,
#         center=[114.8, -11.5],
#         endpoint=[115.4, -7.5],
#         width=[-200, 200],
#         convention="pz",
#         outfile=str(CROSS_SECTION)
#     )

#     fig.shift_origin(yshift=7.4, xshift=17.5)
#     # Buat peta cross section dengan basemap
#     fig.basemap(
#         projection="X4c/-4c",
#         region=[0, 450, 0, 750],
#         frame=['xafg+lDistance (km)','yafg+lDepth (km)', "wsEN"]
#     )

#     # Membuat garis cross section   
#     cross = pd.read_csv(CROSS_SECTION, delimiter='\s+', header=None)
#     # Define your column names
#     header = ['Distance', 'Dept', 'Mag']

#     # Assign the header
#     cross.columns = header

#     fig.plot(x=cross.Distance, 
#                 y=cross.Dept, 
#                 projection="X4/-4", 
#                 size =0.035 * (2 * cross.Mag), style="cc", 
#                 fill = cross.Dept, 
#                 cmap = "quakes.cpt", 
#                 pen = 'black'),
#                 #region = region_cs)
#     fig.text(x=20, y=30, text="A", font="9,Helvetica")
#     fig.text(x=425, y=30, text="A1", font="9,Helvetica")
#     folder_output = os.path.join(OUTPUT_DIR, nama_folder)
#     output_map_filename = os.path.join(folder_output, f'Peta Seismisitas Periode {nama_folder}.png')
#     # Menyimpan gambar
#     fig.savefig(fname=output_map_filename)
#     return nama_folder

def print_report():
    # Set locale to Indonesian for date formatting
    global nama_folder
    locale.setlocale(locale.LC_TIME, 'id_ID.UTF-8')
    folder_output = os.path.join(OUTPUT_DIR, nama_folder)
    
    # Memilih file Excel
    file_path = os.path.join(
        folder_output,
        f"Analisis Gempabumi Periode {nama_folder}.xlsx"
    )
    if not file_path:
        print("Tidak ada file yang dipilih.")
    else:
        # Load the Excel file
        earthquake_data = pd.read_excel(file_path)

        # Pastikan kolom 'Tanggal' dalam format datetime
        earthquake_data['Tanggal'] = pd.to_datetime(earthquake_data['Tanggal'])

        # Path to images
        magnitude_bar_chart_path = os.path.join(folder_output, "diagbat_mag.png")
        magnitude_pie_chart_path = os.path.join(folder_output, "diaglingk_mag.png")

        depth_bar_chart_path = os.path.join(folder_output, "diagbat_depth.png")
        depth_pie_chart_path = os.path.join(folder_output, "diaglingk_depth.png")

        seismic_map_path = os.path.join(
            folder_output,
            f"Peta Seismisitas Periode {nama_folder}.png"
        )
        TEMPLATE_PATH = TEMPLATE_DOC
        earthquake_data = pd.read_excel(file_path)

        # Calculate the required values
        total_earthquakes = len(earthquake_data)
        min_magnitude = earthquake_data['Magnitude'].min()
        max_magnitude = earthquake_data['Magnitude'].max()        

        # Hitung nomor minggu relatif terhadap awal bulan
        start_date = tanggal_awal
        start_of_month = start_date.replace(day=1)
        week_number = (start_date - start_of_month).days // 7 + 1
        bulan_formatted = start_date.strftime('%B')
        end_date =tanggal_akhir

        end_date_formatted = end_date.strftime('%d %B %Y')  
        tanggal_inggris = end_date.strftime('%B, %d')

        # Menghitung jumlah kejadian gempabumi total
        total_earthquakes = len(earthquake_data)

        # Menghitung jumlah gempa dengan lat > -9 dan kedalaman < 60 km
        shallow_earthquake_count = earthquake_data[(earthquake_data['Lintang'] > -9) & 
                                                   (earthquake_data['Kedalaman'] < 60)].shape[0]

        # Menghitung jumlah gempa di selatan Pulau Jawa, Bali, dan Lombok
        southern_earthquake_count = total_earthquakes - shallow_earthquake_count

        # Magnitude categories
        magnitude_below_3 = earthquake_data[earthquake_data['Magnitude'] <= 3]
        magnitude_3_to_5 = earthquake_data[(earthquake_data['Magnitude'] > 3) & (earthquake_data['Magnitude'] < 5)]
        magnitude_above_5 = earthquake_data[earthquake_data['Magnitude'] >= 5]

        count_below_3 = len(magnitude_below_3)
        count_3_to_5 = len(magnitude_3_to_5)
        count_above_5 = len(magnitude_above_5)

        percentage_below_3 = int((count_below_3 / total_earthquakes) * 100)
        percentage_3_to_5 = int((count_3_to_5 / total_earthquakes) * 100)
        percentage_above_5 = int((count_above_5 / total_earthquakes) * 100)

        # Magnitude categories
        depth_below_60 = earthquake_data[earthquake_data['Kedalaman'] <= 60]
        depth_60_to_300 = earthquake_data[(earthquake_data['Kedalaman'] > 60) & (earthquake_data['Kedalaman'] < 300)]
        depth_above_300 = earthquake_data[earthquake_data['Kedalaman'] >= 300]

        count_depth_below_60 = len(depth_below_60)
        count_depth_60_to_300 = len(depth_60_to_300)
        count_depth_above_300 = len(depth_above_300)

        percentage_depth_below_60 = int((count_depth_below_60 / total_earthquakes) * 100)
        percentage_depth_60_to_300 = int((count_depth_60_to_300 / total_earthquakes) * 100)
        percentage_depth_above_300 = int((count_depth_above_300 / total_earthquakes) * 100)

        # Determine dominant depth category
        dominant_category = ''
        if percentage_depth_below_60 >= percentage_depth_60_to_300 and percentage_depth_below_60 >= percentage_depth_above_300:
            dominant_category = f'dangkal (≤60 km) sebesar {percentage_depth_below_60}%'
        elif percentage_depth_60_to_300 >= percentage_depth_below_60 and percentage_depth_60_to_300 >= percentage_depth_above_300:
            dominant_category = f'kedalaman menengah (antara 60 - 300 km) sebesar {percentage_depth_60_to_300}%'
        else:
            dominant_category = f'kedalaman dalam (≥300 km) sebesar {percentage_depth_above_300}%'

        # Menghitung jumlah kejadian gempa yang dirasakan
        earthquakes_felt = earthquake_data[earthquake_data['Dirasakan'].notna()]
        total_felt = len(earthquakes_felt)

        # Create a new Word document
        Laporan_document = Document(TEMPLATE_PATH)

        # Ukuran kertas F4 (21.0 cm x 33.0 cm)
        f4_width = Cm(21.0)
        f4_height = Cm(33.0)

        # Mengatur ukuran kertas F4 untuk setiap section di dokumen
        for section in Laporan_document.sections:
            # Mengatur orientasi potrait untuk F4
            section.orientation = WD_ORIENT.PORTRAIT  # Bisa diubah ke WD_ORIENT.LANDSCAPE jika ingin landscape
            section.page_width = f4_width
            section.page_height = f4_height

            # Mengatur margin halaman (optional, bisa disesuaikan)
            section.top_margin = Cm(2.5)     # Marginn atas
            section.bottom_margin = Cm(2.5)  # Marginn bawah
            section.left_margin = Cm(2.5)    # Marginn kiri
            section.right_margin = Cm(2.5)   # Marginn kanan

        # # Add Title
        title = Laporan_document.add_paragraph()
        title_run = title.add_run(f'STASIUN GEOFISIKA DENPASAR                                                                                              LAPORAN AKTIVITAS SEISMISITAS WILAYAH BALI DAN SEKITARNYA                                                                                              PERIODE {nama_folder}')
        title_run.bold = True
        title_run.font.size = Pt(14)
        title_run.font.all_caps = True  # Membuat semua huruf kapital
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Add space
        Laporan_document.add_paragraph('')

        # Teks Deskriptif Magnitudo dengan perhitungan minggu bulanan
        magnitude_descriptive_text = (
            f"Berdasarkan data Stasiun Geofisika Denpasar selama minggu ke-{week_number} bulan "
            f"{bulan_formatted} {start_date.year}, "
            f"di daerah Bali dan sekitarnya telah terjadi {total_earthquakes} kejadian gempabumi dengan magnitudo bervariasi "
            f"mulai dari M {min_magnitude} sampai M {max_magnitude}. Kejadian gempabumi dengan magnitudo M≤3 sejumlah "
            f"{count_below_3} kejadian atau {percentage_below_3}% dari total kejadian gempabumi. Sedangkan untuk 3<M<5"
            f" sejumlah {count_3_to_5} kejadian atau {percentage_3_to_5}% dari total kejadian gempabumi "
        )

        # Conditional text for magnitude above 5
        if count_above_5 > 0:
            magnitude_descriptive_text += (
                f"dan kejadian gempa M≥5 sejumlah {count_above_5} kejadian atau {percentage_above_5}% dari total kejadian gempabumi."
            )
        else:
            magnitude_descriptive_text += "dan tidak terdapat kejadian gempabumi dengan magnitudo M≥5."

        # Teks Deskriptif Kedalaman
        depth_descriptive_text = (
            f"Berdasarkan kedalaman, gempabumi yang mendominasi adalah kejadian gempabumi {dominant_category}  dari total kejadian gempabumi. "
            f"Jumlah gempabumi dangkal sebanyak {count_depth_below_60} kejadian gempabumi, "
            f"terdapat sebanyak {count_depth_60_to_300} kejadian gempabumi dengan kedalaman menengah antara 60 - 300 km atau {percentage_depth_60_to_300}% dari total kejadian gempabumi, "
        )

        # Conditional text for depth above 300 km
        if count_depth_above_300 > 0:
            depth_descriptive_text += (
                f"dan kejadian gempa dalam sejumlah {count_depth_above_300} kejadian atau {percentage_depth_above_300}% dari total kejadian gempabumi."
            )
        else:
            depth_descriptive_text += "dan tidak terdapat kejadian gempabumi dengan kedalaman dalam ≥ 300 km."

        # Teks analisis yang diminta
        analisis_text1 = (
            f"Selama periode ini, gempabumi yang terjadi di Bali dan sekitarnya dikelompokan berdasarkan sumbernya (Gambar 3):\n"
        )

        analisis_text2 = (
            f"1. Sebanyak {southern_earthquake_count} kejadian gempabumi terjadi di selatan Pulau Jawa, Bali, dan Lombok "
            f"yang diperkirakan berasosiasi dengan subduksi Indo Australia-Eurasia.\n"
            f"2. Sebanyak {shallow_earthquake_count} kejadian gempabumi terjadi menyebar di wilayah Pulau Bali, Lombok dan sekitarnya "
            f"diperkirakan berkaitan dengan sesar di belakang busur Flores atau Flores Back Arc Thrust dan aktivitas sesar aktif."
        )

        # Menyusun narasi pembuka
        if total_felt > 0:
            felt_intro = (
                f"Selama periode ini terdapat {total_felt} kejadian gempabumi yang dirasakan oleh masyarakat "
                f"di wilayah Bali dan sekitarnya."
            )
        else:
            felt_intro = (
                f"Selama periode ini tidak terdapat kejadian gempabumi yang dirasakan oleh masyarakat "
                f"di wilayah Bali dan sekitarnya."
            )

        # Menambahkan paragraf dengan teks deskriptif magnitudo
        magnitude_paragraph = Laporan_document.add_paragraph(magnitude_descriptive_text)
         

        # Add Magnitude Bar Chart Image
        try:
            Laporan_document.add_paragraph().add_run().add_picture(magnitude_bar_chart_path, width=Inches(5.65))
            Laporan_document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except FileNotFoundError:
            Laporan_document.add_paragraph('DIAGRAM BATANG MAGNITUDO TIDAK DITEMUKAN').bold = True

        # Add Magnitude Pie Chart Image
        try:
            Laporan_document.add_paragraph().add_run().add_picture(magnitude_pie_chart_path, width=Inches(3.15))
            Laporan_document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except FileNotFoundError:
            Laporan_document.add_paragraph('DIAGRAM LINGKARAN MAGNITUDO TIDAK DITEMUKAN').bold = True

        # Teks Keterangan Gambar Magnitudo
        Laporan_document.add_paragraph('Gambar 1. Histogram gempabumi harian berdasarkan magnitudo (atas) dan diagram gempabumi berdasarkan magnitudo (bawah)').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Add space
        Laporan_document.add_paragraph('')

        Laporan_depth_paragraph = Laporan_document.add_paragraph(depth_descriptive_text)
        Laporan_depth_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        # Add Depth Bar Chart Image
        try:
            Laporan_document.add_paragraph().add_run().add_picture(depth_bar_chart_path, width=Inches(5.65))
            Laporan_document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except FileNotFoundError:
            Laporan_document.add_paragraph('DIAGRAM BATANG KEDALAMAN TIDAK DITEMUKAN').bold = True

        # Add Depth Pie Chart Image
        try:
            Laporan_document.add_paragraph().add_run().add_picture(depth_pie_chart_path, width=Inches(3.15))
            Laporan_document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except FileNotFoundError:
            Laporan_document.add_paragraph('DIAGRAM LINGKARAN KEDALAMAN TIDAK DITEMUKAN').bold = True

        # Teks Keterangan Gambar Kedalaman (akan diisi sesuai format)
        Laporan_document.add_paragraph('Gambar 2.  Histogram gempabumi harian berdasarkan kedalaman (atas) dan diagram gempabumi berdasarkan kedalaman (bawah)').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Add Seismic Map Image
        try:
            Laporan_document.add_paragraph().add_run().add_picture(seismic_map_path, width=Inches(6))
            Laporan_document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except FileNotFoundError:
            Laporan_document.add_paragraph('PETA SEISMISITAS TIDAK DITEMUKAN').bold = True

        # Teks Keterangan Gambar Peta Seismisitas
        peta_seismisitas_text = (
            f"Gambar 3. Peta seismisitas Bali dan sekitarnya periode {nama_folder}"
        )
        peta_seismisitas_paragraph = Laporan_document.add_paragraph(peta_seismisitas_text)
        peta_seismisitas_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Add space
        Laporan_document.add_paragraph('')

        
        # Menambahkan paragraf analisis ke dalam dokumen
        analisis_paragraph1 = Laporan_document.add_paragraph(analisis_text1)
        analisis_paragraph1.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        # Menambahkan paragraf analisis ke dalam dokumen
        analisis_paragraph2 = Laporan_document.add_paragraph(analisis_text2)
        analisis_paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        # Menambahkan narasi pembuka ke dokumen
        felt_intro_paragraph = Laporan_document.add_paragraph(felt_intro)
        felt_intro_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        # Menyusun detail setiap kejadian gempa yang dirasakan dalam format list dengan nomor urut
        if total_felt > 0:
            # Inisialisasi counter untuk nomor urut
            counter = 1

            for index, row in earthquakes_felt.iterrows():
                magnitude = row['Magnitude']
                date = row['Tanggal'].strftime('%d %B %Y')  # Format tanggal DD/MM/YYYY
                time = row['Waktu (WIB)']  # Asumsikan kolom waktu sudah dalam format string yang benar
                location = row['Dirasakan']
                description = row['Keterangan']
                depth = row['Kedalaman']

                # Format detail gempa yang dirasakan dengan nomor urut
                felt_detail = (
                    f"{counter}. Gempabumi dengan magnitudo {magnitude} {location} yang terjadi pada tanggal {date} "
                    f"pukul {time} WIB berlokasi di {description} pada kedalaman {depth} km."
                )

                # Menambahkan detail gempa ke dalam dokumen sebagai paragraf terpisah
                felt_detail_paragraph = Laporan_document.add_paragraph(felt_detail)
                felt_detail_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                # Meningkatkan counter untuk nomor urut berikutnya
                counter += 1
        else:
            # Jika tidak ada gempa yang dirasakan, tambahkan keterangan tambahan
            no_felt_detail = Laporan_document.add_paragraph(" ")
            no_felt_detail.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


        # Tambahkan section baru dengan orientasi landscape
        new_section = Laporan_document.add_section(WD_SECTION.NEW_PAGE)
        new_section.orientation = WD_ORIENT.LANDSCAPE

        # Atur ukuran kertas agar landscape sesuai
        new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width

        # Menentukan margin untuk section baru agar tabel berada di tengah secara vertikal
        section_margin = Cm(2)  # Sesuaikan margin jika perlu
        new_section.top_margin = section_margin
        new_section.bottom_margin = section_margin
        new_section.left_margin = section_margin
        new_section.right_margin = section_margin

        # Menambahkan paragraf baru untuk membungkus tabel
        table_paragraph = Laporan_document.add_paragraph()
        table_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Mengatur alignment paragraf ke tengah

        # Membuat teks keterangan tabel gempabumi
        tabel_keterangan_text = (
            f"Tabel 1. Data gempabumi di Bali dan sekitarnya tanggal {nama_folder}"
        )

        # Menambahkan teks keterangan tabel ke dalam dokumen
        tabel_keterangan_paragraph = Laporan_document.add_paragraph(tabel_keterangan_text)
        tabel_keterangan_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Mengatur alignment menjadi tengah

        # Menambahkan tabel langsung ke dalam paragraf
        table = Laporan_document.add_table(rows=1, cols=len(earthquake_data.columns))
        table.style = 'Table Grid'

        # Set table alignment to center
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Set column widths to match the text content
        column_widths = {
            "NO.": 0.5,  # NO
            "Tanggal": 2.5,  # TANGGAL
            "Waktu (WIB)": 2.5,  # WAKTU (WIB)
            "Lintang": 2.0,  # LATITUDE
            "Bujur": 2.0,  # LONGITUDE
            "Kedalaman": 2.0,  # Depth
            "Magnitude": 1.5,  # Magnitude
            "Keterangan": 6.0,  # Keterangan
            "Dirasakan": 2.5  # Dirasakan
        }

        # Add header row
        hdr_cells = table.rows[0].cells
        for i, column_name in enumerate(earthquake_data.columns):
            hdr_cells[i].text = column_name
            hdr_cells[i].width = Cm(column_widths.get(column_name, 2.0))  # Set width based on column

        # Add data rows, memformat kolom 'Tanggal' tanpa jam
        for index, row in earthquake_data.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                # Jika kolom adalah 'Tanggal', hanya tampilkan tanggal tanpa jam
                if earthquake_data.columns[i] == 'Tanggal':
                    row_cells[i].text = value.strftime('%d/%m/%Y')  # Format tanggal menjadi DD/MM/YYYY
                # Jika kolom adalah 'Dirasakan' dan nilai adalah NaN, ganti dengan teks yang sesuai
                elif earthquake_data.columns[i] == 'Dirasakan':
                    # Ganti NaN dengan string kosong atau teks lain
                    row_cells[i].text = '' if pd.isna(value) else str(value)
                else:
                    row_cells[i].text = str(value)

                # Set the width of each data cell, sesuai dengan kolom
                row_cells[i].width = Cm(column_widths.get(earthquake_data.columns[i], 2.0))

        # Mengatur tabel agar berada di tengah secara horizontal
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Simpan 'datalengkap.csv' ke folder baru dengan nama 'data gempa (nama folder).xlsx'
        Laporan_output_file = os.path.join(folder_output, f'Laporan Gempabumi Periode {nama_folder}.docx')
    
        # Save the document
        Laporan_document.save(Laporan_output_file)

        Release_document=Document(TEMPLATE_PATH)

        # Mengatur ukuran kertas F4 untuk setiap section di dokumen
        for section in Release_document.sections:
            # Mengatur orientasi potrait untuk F4
            section.orientation = WD_ORIENT.PORTRAIT  # Bisa diubah ke WD_ORIENT.LANDSCAPE jika ingin landscape
            section.page_width = f4_width
            section.page_height = f4_height

            # Mengatur margin halaman (optional, bisa disesuaikan)
            section.top_margin = Cm(2.5)     # Marginn atas
            section.bottom_margin = Cm(2.5)  # Marginn bawah
            section.left_margin = Cm(2.5)    # Marginn kiri
            section.right_margin = Cm(2.5)   # Marginn kanan

        # Add Title
        title = Release_document.add_paragraph()
        title_run = title.add_run(f'PRESS RELEASE AKTIVITAS SEISMISITAS                                                                                              WILAYAH BALI DAN SEKITARNYA                                                                                                        PERIODE {nama_folder}')
        title_run.bold = True
        title_run.font.size = Pt(14)
        title_run.font.all_caps = True  # Membuat semua huruf kapital
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Add space
        Release_document.add_paragraph('')

        Release_magnitude_paragraph = Release_document.add_paragraph(magnitude_descriptive_text)
        Release_magnitude_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        Release_depth_paragraph = Release_document.add_paragraph(depth_descriptive_text)
        Release_depth_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        Release_document.add_paragraph(analisis_text1)
        Release_document.add_paragraph(analisis_text2)
        Release_document.add_paragraph(felt_intro)

        # Menyusun detail setiap kejadian gempa yang dirasakan dalam format list dengan nomor urut
        if total_felt > 0:
            # Inisialisasi counter untuk nomor urut
            counter = 1

            for index, row in earthquakes_felt.iterrows():
                magnitude = row['Magnitude']
                date = row['Tanggal'].strftime('%d %B %Y')  # Format tanggal DD/MM/YYYY
                time = row['Waktu (WIB)']  # Asumsikan kolom waktu sudah dalam format string yang benar
                location = row['Dirasakan']
                description = row['Keterangan']
                depth = row['Kedalaman']

                # Format detail gempa yang dirasakan dengan nomor urut
                felt_detail = (
                    f"{counter}. Gempabumi dengan magnitudo {magnitude} {location} yang terjadi pada tanggal {date} "
                    f"pukul {time} WIB berlokasi di {description} pada kedalaman {depth} km."
                )

                # Menambahkan detail gempa ke dalam dokumen sebagai paragraf terpisah
                felt_detail_paragraph = Release_document.add_paragraph(felt_detail)
                felt_detail_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                # Meningkatkan counter untuk nomor urut berikutnya
                counter += 1
        else:
            # Jika tidak ada gempa yang dirasakan, tambahkan keterangan tambahan
            no_felt_detail = Release_document.add_paragraph(" ")
            no_felt_detail.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        Release_document.add_page_break()

         # Add Magnitude Bar Chart Image
        try:
            Release_document.add_paragraph().add_run().add_picture(magnitude_bar_chart_path, width=Inches(5.65))
            Release_document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except FileNotFoundError:
            Release_document.add_paragraph('DIAGRAM BATANG MAGNITUDO TIDAK DITEMUKAN').bold = True

        # Add Magnitude Pie Chart Image
        try:
            Release_document.add_paragraph().add_run().add_picture(magnitude_pie_chart_path, width=Inches(3.15))
            Release_document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except FileNotFoundError:
            Release_document.add_paragraph('DIAGRAM LINGKARAN MAGNITUDO TIDAK DITEMUKAN').bold = True

        # Teks Keterangan Gambar Magnitudo
        Release_document.add_paragraph('Gambar 1. Histogram gempabumi harian berdasarkan magnitudo (atas) dan diagram gempabumi berdasarkan magnitudo (bawah)').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        Release_document.add_page_break()
        # Add Depth Bar Chart Image
        try:
            Release_document.add_paragraph().add_run().add_picture(depth_bar_chart_path, width=Inches(5.65))
            Release_document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except FileNotFoundError:
            Release_document.add_paragraph('DIAGRAM BATANG KEDALAMAN TIDAK DITEMUKAN').bold = True

        # Add Depth Pie Chart Image
        try:
            Release_document.add_paragraph().add_run().add_picture(depth_pie_chart_path, width=Inches(3.15))
            Release_document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except FileNotFoundError:
            Release_document.add_paragraph('DIAGRAM LINGKARAN KEDALAMAN TIDAK DITEMUKAN').bold = True

        # Teks Keterangan Gambar Kedalaman (akan diisi sesuai format)
        Release_document.add_paragraph('Gambar 2.  Histogram gempabumi harian berdasarkan kedalaman (atas) dan diagram gempabumi berdasarkan kedalaman (bawah)').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        Release_document.add_page_break()
        # Add Seismic Map Image
        try:
            Release_document.add_paragraph().add_run().add_picture(seismic_map_path, width=Inches(6))
            Release_document.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except FileNotFoundError:
            Release_document.add_paragraph('PETA SEISMISITAS TIDAK DITEMUKAN').bold = True

        # Teks Keterangan Gambar Peta Seismisitas
        peta_seismisitas_text = (
            f"Gambar 3. Peta seismisitas Bali dan sekitarnya periode {nama_folder}"
        )
        peta_seismisitas_paragraph = Release_document.add_paragraph(peta_seismisitas_text)
        peta_seismisitas_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Tambahkan section baru dengan orientasi landscape
        new_section = Release_document.add_section(WD_SECTION.NEW_PAGE)
        new_section.orientation = WD_ORIENT.LANDSCAPE

        # Atur ukuran kertas agar landscape sesuai
        new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width

        # Menambahkan paragraf baru untuk membungkus tabel
        table_paragraph = Release_document.add_paragraph()
        table_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Mengatur alignment paragraf ke tengah

        # Membuat teks keterangan tabel gempabumi
        tabel_keterangan_text = (
            f"Tabel 1. Data gempabumi di Bali dan sekitarnya tanggal {nama_folder}"
        )

        # Menambahkan teks keterangan tabel ke dalam dokumen
        tabel_keterangan_paragraph = Release_document.add_paragraph(tabel_keterangan_text)
        tabel_keterangan_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Mengatur alignment menjadi tengah

        # Menambahkan tabel langsung ke dalam paragraf
        table = Release_document.add_table(rows=1, cols=len(earthquake_data.columns))
        table.style = 'Table Grid'

        # Set table alignment to center
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Set column widths to match the text content
        column_widths = {
            "NO.": 0.5,  # NO
            "Tanggal": 2.5,  # TANGGAL
            "Waktu (WIB)": 2.5,  # WAKTU (WIB)
            "Lintang": 2.0,  # LATITUDE
            "Bujur": 2.0,  # LONGITUDE
            "Kedalaman": 2.0,  # Depth
            "Magnitude": 1.5,  # Magnitude
            "Keterangan": 6.0,  # Keterangan
            "Dirasakan": 2.5  # Dirasakan
        }

        # Add header row
        hdr_cells = table.rows[0].cells
        for i, column_name in enumerate(earthquake_data.columns):
            hdr_cells[i].text = column_name
            hdr_cells[i].width = Cm(column_widths.get(column_name, 2.0))  # Set width based on column

        # Add data rows, memformat kolom 'Tanggal' tanpa jam
        for index, row in earthquake_data.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                # Jika kolom adalah 'Tanggal', hanya tampilkan tanggal tanpa jam
                if earthquake_data.columns[i] == 'Tanggal':
                    row_cells[i].text = value.strftime('%d/%m/%Y')  # Format tanggal menjadi DD/MM/YYYY
                # Jika kolom adalah 'Dirasakan' dan nilai adalah NaN, ganti dengan teks yang sesuai
                elif earthquake_data.columns[i] == 'Dirasakan':
                    # Ganti NaN dengan string kosong atau teks lain
                    row_cells[i].text = '' if pd.isna(value) else str(value)
                else:
                    row_cells[i].text = str(value)

                # Set the width of each data cell, sesuai dengan kolom
                row_cells[i].width = Cm(column_widths.get(earthquake_data.columns[i], 2.0))

        # Mengatur tabel agar berada di tengah secara horizontal
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Simpan 'datalengkap.csv' ke folder baru dengan nama 'data gempa (nama folder).xlsx'
        Release_output_file = os.path.join(folder_output, f'Release Gempabumi Periode {nama_folder}.docx')
    
        # Save the document
        Release_document.save(Release_output_file)

        Caption_document=Document()
        Caption_document.add_paragraph("###Caption Ig dan Fb (Jangan lupa ganti tanggalnya!)")
        Caption_document.add_paragraph("Halo sobat BMKG Denpasar! Apa kabar? Semoga sehat ya🤗\n" 
                            f"Berikut kami informasikan aktivitas kegempaan wilayah Bali dan sekitarnya periode {nama_folder}."
                            "Pada minggu ini, aktivitas kegempaan didominasi oleh gempa magnitudo M < 3 dengan kedalaman dangkal < 60 km."
                            f"Pada minggu ini terdapat {total_earthquakes} kejadian gempabumi." \
                            "\n" \
                            "\n" \
                            f"Jumat, {end_date_formatted}.\n"
                            ".\n"
                            ".\n"
                            ".\n"
                            "Hello BMKG Denpasar friends! How are you? Hope you are healthy🤗\n" 
                            f"Here we provide information on earthquake activity in the Bali area and its surroundings for the period {nama_folder}."
                            "This week, seismic activity was dominated by earthquakes with a magnitude of M < 3 with a shallow depth of < 60 km."
                            f"This week there are {total_earthquakes} earthquake event. Greetings of health and enthusiasm." \
                            "\n" \
                            "\n" \
                            f"{tanggal_inggris}.\n"
                            "#stageofdenpasar #SeismisitasBali #GempaBali #GempaBumi #Seismologi #AktivitasSeismik #PetaGempa #DataGempa #SeismisitasIndonesia #akuntabel #siapuntukselamat ")

        # Add space
        Caption_document.add_paragraph('')
        Caption_document.add_paragraph("###Caption X")
        Caption_document.add_paragraph("Halo sobat BMKG Denpasar! Apa kabar? Semoga sehat ya🤗.\n"
                            f"Berikut kami informasikan aktivitas kegempaan wilayah Bali dan sekitarnya periode {nama_folder}.\n"
                            "Salam sehat dan semangat. \n"
                            "#stageofdenpasar #SeismisitasBali #GempaBali #GempaBumi #Seismologi #AktivitasSeismik")
        print(type(count_3_to_5))
        print(type(count_below_3))
        print(type(depth_60_to_300))
        print(type(depth_below_60))
        print(type(total_felt))

        if total_felt > 0 :
            teks_dirasakan=(
                f"Stasiun Geofisika Denpasar mencatat {total_earthquakes} kejadian gempabumi dan terdapat {total_felt} kejadian gempabumi dirasakan di wilayah bali dan sekitarnya"
            )
        else:
            teks_dirasakan=(
                f"Stasiun Geofisika Denpasar mencatat {total_earthquakes} kejadian gempabumi dan tidak terdapat kejadian gempabumi dirasakan di wilayah bali dan sekitarnya"
            )

        if count_3_to_5 > count_below_3:
            teks_mag=(
                f"pada minggu ini di dominasi gempabumi magnitudo menengah 3-5 dengan total {count_3_to_5} kejadian gempabumi "
                f"dan {count_below_3} kejadian gempabumi kecil dibawah 3 magnitudo dengan magnitudo maksimal {max_magnitude} magnitudo"
            )
        else:
            teks_mag=(
                f"pada minggu ini di dominasi gempabumi magnitudo kecil dibawah 3 magnitudo dengan total {count_below_3} kejadian gempabumi "
                f"dan {count_3_to_5} kejadian gempabumi menengah dengan magnitudo maksimal {max_magnitude} magnitudo"
            )

        if count_depth_60_to_300 > count_depth_below_60:
            teks_kedalaman=(
                f"Berdasarkan kedalaman didominasi gempabumi menengah 60 sampai 300 km dengan total {count_depth_60_to_300} kejadian gempabumi"
                f"dan {count_depth_below_60} kejadian gempabumi dangkal kurang dari 60 km"
            )
        else:
            teks_kedalaman=(
                f"Berdasarkan kedalaman didominasi gempabumi dangkal kurang dari 60 km dengan total {count_depth_below_60} kejadian gempabumi"
                f"dan gempabumi kedalaman menengah {count_depth_60_to_300} kejadian"
            )
        Caption_document.add_paragraph('')
        Caption_document.add_paragraph("###Teks untuk voicenote(untuk wilayah dirasakan sebutkan di wilayah bali aja)")
        Caption_document.add_paragraph(f"Rangkuman Aktivitas Seismisitas Wilayah Bali dan Sekitarnya peiode {nama_folder}")
        Caption_document.add_paragraph(teks_dirasakan)
        Caption_document.add_paragraph(teks_mag)
        Caption_document.add_paragraph(teks_kedalaman)

        # Menyusun detail setiap kejadian gempa yang dirasakan dalam format list dengan nomor urut
        if total_felt > 0:
            # Inisialisasi counter untuk nomor urut
            counter = 1

            for index, row in earthquakes_felt.iterrows():
                magnitude = row['Magnitude']
                date = row['Tanggal'].strftime('%d %B %Y')  # Format tanggal DD/MM/YYYY
                time = row['Waktu (WIB)']  # Asumsikan kolom waktu sudah dalam format string yang benar
                location = row['Dirasakan']
                description = row['Keterangan']
                depth = row['Kedalaman']

                # Format detail gempa yang dirasakan dengan nomor urut
                felt_detail = (
                    f"{counter}. Gempabumi pada tanggal {date} dengan magnitudo {magnitude} {location}"
                    f" dengan pusat gempa berlokasi di {description} pada kedalaman {depth} km."
                )

                # Menambahkan detail gempa ke dalam dokumen sebagai paragraf terpisah
                felt_detail_paragraph = Caption_document.add_paragraph(felt_detail)
                felt_detail_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                # Meningkatkan counter untuk nomor urut berikutnya
                counter += 1
        else:
            # Jika tidak ada gempa yang dirasakan, tambahkan keterangan tambahan
            no_felt_detail = Caption_document.add_paragraph(" ")
            no_felt_detail.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        Caption_document.add_paragraph("Demikian disampaikan, terima kasih.")

        Caption_document.add_page_break()
        # Add space
        Caption_document.add_paragraph('')

        # Tambah penerima
        Caption_document.add_paragraph("Yth.\n1. Ibu Deputi Bidang Geofisika\n"
                            "2. Bapak Direktur Gempabumi dan Tsunami\n"
                            "3. Ibu Direktur Seismologi Teknik, Geofisika Potensial dan Tanda Waktu\n"
                            "4. Bapak Kepala Balai Wilayah III\n")

        # Tambah salam pembuka
        Caption_document.add_paragraph("Dengan hormat,\nBerikut kami sampaikan sebagai berikut:")

        # Tambah poin pertama
        Caption_document.add_paragraph(f"Infografis seismisitas wilayah Bali periode {nama_folder}. ")

        # Tambah daftar link media sosial
        Caption_document.add_paragraph("\nBerikut terlampir link media sosial infografis dan videografis seismisitas wilayah Bali dan sekitarnya dari Stasiun Geofisika Denpasar:")

        social_links = [
            ("1. Link Instagram", "Videografis dan Infografis: "),
            ("2. Link Facebook", "Videografis dan Infografis: "),
            ("3. Link X", "Videografis dan Infografis: "),
            ("4. Link berita website", ": https://stageof-bali.bmkg.go.id/informasi-mingguan"),
        ]

        for title, link in social_links:
            p = Caption_document.add_paragraph()
            p.add_run(title).bold = True
            p.add_run("\n" + link)

        # Tambah penutup
        Caption_document.add_paragraph("\nDemikian, mohon berkenan dan terima kasih.")
        Caption_document.add_paragraph("Hormat kami,\n \nBMKG Stasiun Geofisika Denpasar")

        # Tambah kontak
        Caption_document.add_paragraph("\nKontak Stasiun Geofisika Denpasar:")
        contacts = [
            "📍 Alamat: Jl. Pulau Tarakan No. 1 Sanglah, Denpasar-Bali",
            "📞 Telepon: (0361) 226157",
            "🌐 Web: stageof-bali.bmkg.go.id",
            "📧 Email: stageof.denpasar@bmkg.go.id"
        ]

        for contact in contacts:
            Caption_document.add_paragraph(contact)

        Caption_document.add_section()

        # Tambah kalimat penutup
        Caption_document.add_paragraph()
        Caption_document.add_paragraph("Om Swastiastu, Selamat Pagi sobat BMKG. Berikut kami sampaikan informasi geofisika dari Stasiun Geofisika Denpasar yang dapat diakses melalui link berikut:\n" 
                            "\n https://linktr.ee/StageofDenpasar \n"
                            "\n Informasi meliputi:"
                            "\n ✔ Informasi Sambaran Petir Mingguan"
                            "\n ✔ Informasi Waktu Terbit dan Tenggelam Matahari Mingguan"
                            "\n ✔ Informasi Gempabumi Mingguan \n"
                            "\n Demikian, mohon berkenan dan terima kasih. \n"
                            "\n Hormat kami,\n BMKG Stasiun Geofisika Denpasar"
                            " \n"
                            " \n"
                            "\nKontak Stasiun Geofisika Denpasar:")
        # Tambah kontak
        contacts = [
            "📍 Alamat: Jl. Pulau Tarakan No. 1 Sanglah, Denpasar-Bali",
            "📞 Telepon: (0361) 226157",
            "🌐 Web: stageof-bali.bmkg.go.id",
            "📧 Email: stageof.denpasar@bmkg.go.id"
        ]

        for contact in contacts:
            Caption_document.add_paragraph(contact)

        Caption_output_file = os.path.join(folder_output, f'Caption Sosmed Periode {nama_folder}.docx')
    
        # Save the document
        Caption_document.save(Caption_output_file)


    #messagebox.showinfo("Cetak Laporan", "Berhasil Mencetak laporan ke file Word")

def sekali_jalan():
    process_data()
    print_analysis()
    # print_map()
    print_report()
    messagebox.showinfo("Seismioto", "Well done mate")