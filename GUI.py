import os
import customtkinter as ctk
from tkinter import filedialog
from tkinter import messagebox
from tkcalendar import DateEntry
import sys
import seismioto_fdsn_engine as engine
import threading
from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent
sys.path.append(str(BASE_DIR))

# ==================================================
# KONFIGURASI
# ==================================================

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PETUNJUK_FILE = os.path.join(BASE_DIR, "petunjuk.txt")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# ==================================================
# WINDOW
# ==================================================

app = ctk.CTk()
app.title("SEISMIOTO V2")
app.geometry("1600x1000")
app.minsize(1200, 700)

# ==================================================
# VARIABLE
# ==================================================



# ==================================================
# FUNCTION
# ==================================================

selected_files = []

def pilih_file():

    global selected_files

    selected_files = filedialog.askopenfilenames(
        title="Pilih XML",
        filetypes=[("XML", "*.xml")]
    )

    if selected_files:
        entry_file.delete(0,"end")
        entry_file.insert(
            0,
            f"{len(selected_files)} file dipilih"
        )

def run_engine():

    try:

        status_label.configure(
            text="Memproses..."
        )

        if not selected_files:
            messagebox.showerror(
                "Error",
                "Pilih file XML terlebih dahulu"
            )

            return

        engine.set_input(
            selected_files,
            awal_entry.get(),
            akhir_entry.get()
        )

        tulis_log("Memulai proses")

        engine.sekali_jalan()

        tulis_log("Selesai")

        status_label.configure(
            text="Analisis selesai"
        )

    except Exception as e:

        tulis_log(
            f"ERROR : {e}"
        )

        status_label.configure(
            text="Gagal"
        )

def jalankan_analisis():

    threading.Thread(
        target=run_engine,
        daemon=True
    ).start()


def buka_output():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    os.startfile(OUTPUT_DIR)


def tulis_log(teks):
    log_box.configure(state="normal")
    log_box.insert("end", teks + "\n")
    log_box.see("end")
    log_box.configure(state="disabled")


def load_petunjuk():

    try:
        with open(
            PETUNJUK_FILE,
            "r",
            encoding="utf-8"
        ) as f:

            isi = f.read()

    except Exception:
        isi = "File petunjuk.txt tidak ditemukan."

    petunjuk_box.insert("1.0", isi)
    petunjuk_box.configure(state="disabled")

# ==================================================
# LAYOUT
# ==================================================

main_frame = ctk.CTkFrame(app)
main_frame.pack(fill="both", expand=True, padx=10, pady=10)

main_frame.grid_columnconfigure(0, weight=3)
main_frame.grid_columnconfigure(1, weight=2)

main_frame.grid_rowconfigure(0, weight=1)

# ==================================================
# PANEL KIRI
# ==================================================

left_frame = ctk.CTkFrame(main_frame)
left_frame.grid(row=0, column=0, sticky="nsew", padx=(0,10))

# --------------------------------------------------

title_label = ctk.CTkLabel(
    left_frame,
    text="SEISMIOTO V2",
    font=("Segoe UI", 28, "bold")
)
title_label.pack(pady=(20,5))

subtitle_label = ctk.CTkLabel(
    left_frame,
    text="Sistem Otomatisasi Informasi Gempabumi",
    font=("Segoe UI", 14)
)
subtitle_label.pack(pady=(0,20))

# ==================================================
# FILE XML
# ==================================================

file_frame = ctk.CTkFrame(left_frame)
file_frame.pack(fill="x", padx=20, pady=10)

ctk.CTkLabel(
    file_frame,
    text="File XML"
).pack(anchor="w", padx=10, pady=(10,5))

entry_file = ctk.CTkEntry(
    file_frame,
    height=35
)
entry_file.pack(
    side="left",
    fill="x",
    expand=True,
    padx=(10,5),
    pady=(0,10)
)

ctk.CTkButton(
    file_frame,
    text="Browse",
    width=120,
    command=pilih_file
).pack(
    side="right",
    padx=(5,10),
    pady=(0,10)
)

# ==================================================
# TANGGAL
# ==================================================

date_frame = ctk.CTkFrame(left_frame)
date_frame.pack(fill="x", padx=20, pady=10)

ctk.CTkLabel(
    date_frame,
    text="Tanggal Awal"
).grid(row=0, column=0, padx=10, pady=10)

awal_entry = DateEntry(
    date_frame,
    date_pattern="dd/mm/yyyy"
)
awal_entry.grid(row=0, column=1, padx=10, pady=10)

ctk.CTkLabel(
    date_frame,
    text="Tanggal Akhir"
).grid(row=1, column=0, padx=10, pady=10)

akhir_entry = DateEntry(
    date_frame,
    date_pattern="dd/mm/yyyy"
)
akhir_entry.grid(row=1, column=1, padx=10, pady=10)

# ==================================================
# BUTTON
# ==================================================

button_frame = ctk.CTkFrame(left_frame)
button_frame.pack(fill="x", padx=20, pady=10)

ctk.CTkButton(
    button_frame,
    text="JALANKAN ANALISIS",
    height=45,
    command=jalankan_analisis
).pack(fill="x", padx=10, pady=(10,5))

ctk.CTkButton(
    button_frame,
    text="BUKA FOLDER OUTPUT",
    height=40,
    command=buka_output
).pack(fill="x", padx=10, pady=(5,10))

# ==================================================
# STATUS
# ==================================================

status_frame = ctk.CTkFrame(left_frame)
status_frame.pack(fill="x", padx=20, pady=10)

ctk.CTkLabel(
    status_frame,
    text="Status"
).pack(anchor="w", padx=10, pady=(10,5))

status_label = ctk.CTkLabel(
    status_frame,
    text="Siap digunakan"
)
status_label.pack(anchor="w", padx=10, pady=(0,10))

# ==================================================
# PROGRESS
# ==================================================

progress_frame = ctk.CTkFrame(left_frame)
progress_frame.pack(fill="x", padx=20, pady=10)

progress = ctk.DoubleVar()

progressbar = ctk.CTkProgressBar(
    progress_frame
)

progressbar.pack(
    fill="x",
    padx=10,
    pady=15
)

progressbar.set(0)

# ==================================================
# LOG
# ==================================================

log_frame = ctk.CTkFrame(left_frame)
log_frame.pack(
    fill="both",
    expand=True,
    padx=20,
    pady=(10,20)
)

ctk.CTkLabel(
    log_frame,
    text="Log Proses"
).pack(anchor="w", padx=10, pady=10)

log_box = ctk.CTkTextbox(
    log_frame
)

log_box.pack(
    fill="both",
    expand=True,
    padx=10,
    pady=(0,10)
)

log_box.configure(state="disabled")

# ==================================================
# PANEL KANAN
# ==================================================

right_frame = ctk.CTkFrame(main_frame)
right_frame.grid(row=0, column=1, sticky="nsew")

ctk.CTkLabel(
    right_frame,
    text="PETUNJUK PENGGUNAAN",
    font=("Segoe UI", 18, "bold")
).pack(pady=15)

petunjuk_box = ctk.CTkTextbox(
    right_frame
)

petunjuk_box.pack(
    fill="both",
    expand=True,
    padx=10,
    pady=(0,10)
)

load_petunjuk()

# ==================================================
# START
# ==================================================

app.mainloop()