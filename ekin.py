# -*- coding: utf-8 -*-
"""
Query bulanan:
- Menit Kirim diambil dari JSON Node-RED (dashboard):

    http://172.19.2.185:5000/origin/REG3/<AGENCY>?bulan=YYYYMM

AGENCY dipakai:
  - BMKG-BALI
  - BMKG-DNP
  - BMKG-MTRM

Dipilih Menit Kirim tercepat (dalam MENIT).

Parameter lain (RMS, ERH, ERZ, M.dist, dll)
diambil dari FDSN final only:

  http://172.19.2.185:18080/fdsnws/event/1/query
    ?eventid=XXX
    &estatus=final
    &nodata=404
    &format=scml

Info Gempa menggunakan jarak + kota terdekat (kabkota.txt).
"""

import os
import math
import time
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Any, List, Optional

import requests
import pandas as pd
from xml.etree import ElementTree as ET
from geopy.distance import geodesic

# ================= CONFIG =================

BASE_ORIGIN_URL = "http://172.19.2.185:5000/origin"
FDSN_URL = "http://172.19.2.185:18080/fdsnws/event/1/query"

REGION_CODE = "REG3"
AGENCIES = ["BMKG-BALI", "BMKG-DNP", "BMKG-MTRM"]

MAX_WORKERS = 8
HTTP_TIMEOUT = 20

# ================= HELPER XML =================

def _local(tag: str) -> str:
    return tag.split("}", 1)[-1]

def _child(el: ET.Element, name: str) -> Optional[ET.Element]:
    for c in el:
        if _local(c.tag) == name:
            return c
    return None

def _children(el: ET.Element, name: str) -> List[ET.Element]:
    return [c for c in el if _local(c.tag) == name]

def parse_iso(s: str) -> Optional[datetime]:
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        return dt.astimezone(timezone.utc)
    except Exception:
        return None

# ================= Kota terdekat =================

CITY_FILE = "kabkota.txt"

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

# ================= JSON ORIGIN FETCH =================

def fetch_json(reg: str, agency: str, yyyymm: str) -> List[Dict[str, Any]]:
    """
    Ambil JSON dari Node-RED:
      /origin/<REG>/<AGENCY>?bulan=YYYYMM

    Struktur contoh:
    {
      "agency": "BMKG-BALI",
      "bulan": "202511",
      "data": [
         { "event_id": "...", "waktu_kirim": 123.45, ... },
         ...
      ]
    }

    Fungsi ini selalu mengembalikan LIST berisi dict event.
    """
    url = f"{BASE_ORIGIN_URL}/{reg}/{agency}"
    params = {"bulan": yyyymm}
    try:
        r = requests.get(url, params=params, timeout=HTTP_TIMEOUT)
        r.raise_for_status()
        obj = r.json()
    except Exception as e:
        print(f"[WARN] gagal ambil JSON dari {agency}: {e}")
        return []

    # Normalisasi → list of events
    if isinstance(obj, dict):
        data = obj.get("data", [])
        if isinstance(data, list):
            return data
        return []
    elif isinstance(obj, list):
        return obj
    else:
        return []

def collect_min_kirim(reg: str, yyyymm: str) -> Dict[str, float]:
    """
    Mengumpulkan Menit Kirim tercepat dari beberapa agency.

    - Ambil data dari:
        BMKG-BALI, BMKG-DNP, BMKG-MTRM
    - Field dipakai:
        event_id      → ID event
        waktu_kirim   → detik
    - Disimpan sebagai MENIT (waktu_kirim / 60).
    """
    store: Dict[str, float] = {}

    for ag in AGENCIES:
        rows = fetch_json(reg, ag, yyyymm)
        for row in rows:
            # adapt ke berbagai kemungkinan nama key
            ev = (
                row.get("event_id")
                or row.get("eventid")
                or row.get("id")
            )
            if not ev:
                continue

            w = (
                row.get("waktu_kirim")
                or row.get("menit_kirim")  # kalau suatu saat diganti nama
            )
            if w is None:
                continue

            try:
                w = float(w)
            except Exception:
                continue

            # JSON pakai DETIK → konversi ke MENIT
            menit = w / 60.0

            if ev not in store:
                store[ev] = menit
            else:
                if menit < store[ev]:
                    store[ev] = menit

    return store

# ================= FDSN FINAL PARSER =================

def fetch_fdsn_final(event_id: str) -> Optional[bytes]:
    params = {
        "eventid": event_id,
        "estatus": "final",
        "nodata": "404",
        "format": "scml",
    }
    try:
        r = requests.get(FDSN_URL, params=params, timeout=HTTP_TIMEOUT)
        r.raise_for_status()
        return r.content
    except Exception as e:
        print(f"[WARN] gagal FDSN final {event_id}: {e}")
        return None

def parse_fdsn_final(xml_bytes: bytes) -> Dict[str, Any]:
    out: Dict[str, Any] = {
        "latitude": None,
        "longitude": None,
        "depth": None,
        "mag": None,
        "origin_time": None,
        "RMS": None,
        "Gap": None,
        "Phase": None,
        "ERH": None,
        "ERZ": None,
        "Mdist": None,
    }
    try:
        root = ET.fromstring(xml_bytes)
    except Exception:
        return out

    ep = _child(root, "EventParameters")
    if ep is None:
        return out

    ev = _child(ep, "event")
    origins = _children(ep, "origin")
    if not origins:
        return out

    # preferred origin
    pref = None
    poi = _child(ev, "preferredOriginID") if ev is not None else None
    prefid = poi.text.strip() if (poi is not None and poi.text) else None
    if prefid:
        for o in origins:
            if o.get("publicID") == prefid:
                pref = o
                break
    if pref is None:
        pref = origins[0]

    # origin time
    ot = _child(pref, "time")
    if ot is not None:
        v = _child(ot, "value")
        if v is not None and v.text:
            out["origin_time"] = parse_iso(v.text)

    # lat lon depth
    latn = _child(pref, "latitude")
    if latn is not None:
        v = _child(latn, "value")
        if v is not None and v.text:
            out["latitude"] = float(v.text)

    lonn = _child(pref, "longitude")
    if lonn is not None:
        v = _child(lonn, "value")
        if v is not None and v.text:
            out["longitude"] = float(v.text)

    dep = _child(pref, "depth")
    if dep is not None:
        v = _child(dep, "value")
        if v is not None and v.text:
            val = float(v.text)
            out["depth"] = val / 1000.0 if val > 100 else val

    # ========== PATCH MAG FROM ORIGIN (BUKAN preferredMagnitude) ==========
    mag_val: Optional[float] = None
    mags_o = _children(pref, "magnitude")  # magnitude di dalam origin

    for m in mags_o:
        v_node = None

        # Skema 1: <magnitude><magnitude><value>...</value></magnitude></magnitude>
        inner_mag = _child(m, "magnitude")
        if inner_mag is not None:
            v_node = _child(inner_mag, "value")

        # Skema 2: <magnitude><mag><value>...</value></mag></magnitude>
        if v_node is None:
            inner_mag = _child(m, "mag")
            if inner_mag is not None:
                v_node = _child(inner_mag, "value")

        if v_node is not None and v_node.text:
            try:
                mag_val = float(v_node.text.strip())
                break  # pakai magnitude pertama yang valid
            except Exception:
                continue

    out["mag"] = mag_val
    # =====================================================================

    # quality
    q = _child(pref, "quality")
    if q is not None:
        v = _child(q, "standardError")
        if v is not None and v.text:
            out["RMS"] = float(v.text)

        v = _child(q, "azimuthalGap")
        if v is not None and v.text:
            out["Gap"] = float(v.text)

        v = _child(q, "usedPhaseCount")
        if v is not None and v.text:
            out["Phase"] = int(float(v.text))

        v = _child(q, "minimumDistance")
        if v is not None and v.text:
            out["Mdist"] = float(v.text)

    # ERH dari uncertainty
    unc = _child(pref, "uncertainty")
    if unc is not None:
        v = _child(unc, "maxHorizontalUncertainty")
        if v is not None and v.text:
            val = float(v.text)
            out["ERH"] = val / 1000.0 if val > 100 else val
        v = _child(unc, "minHorizontalUncertainty")
        if v is not None and v.text and out["ERH"] is None:
            val = float(v.text)
            out["ERH"] = val / 1000.0 if val > 100 else val

    # ERZ dari depth/uncertainty
    if dep is not None:
        v = _child(dep, "uncertainty")
        if v is not None and v.text:
            val = float(v.text)
            out["ERZ"] = val / 1000.0 if val > 100 else val

    return out

# ================= Info Gempa =================

BULAN_INA = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}

def build_info(row: pd.Series) -> str:
    ot: Optional[datetime] = row.get("origin_time")
    lat = row.get("latitude")
    lon = row.get("longitude")
    dep = row.get("depth")
    mag = row.get("mag")

    if ot is None:
        tgl_str = "-"
    else:
        dt2 = ot + timedelta(hours=7)
        tgl_str = f"{dt2.day:02d}-{BULAN_INA[dt2.month]}-{dt2.year} {dt2.strftime('%H:%M:%S')} WIB"

    if pd.isna(lat) or pd.isna(lon):
        lat_s = "NA"
        lon_s = "NA"
        lok_setempat = "-"
    else:
        lat = float(lat)
        lon = float(lon)
        lat_s = f"{abs(lat):.2f} {'LS' if lat < 0 else 'LU'}"
        lon_s = f"{abs(lon):.2f} {'BT' if lon > 0 else 'BB'}"
        lok_setempat = nearest_city(lat, lon)

    depth_km = 0 if pd.isna(dep) else int(round(float(dep)))
    mag_s = "NA" if pd.isna(mag) else f"{float(mag):.1f}"

    return (
        f"Info Gempa Mag:{mag_s}, {tgl_str}, "
        f"Lok: {lat_s} - {lon_s} ({lok_setempat}), "
        f"Kedalaman: {depth_km} Km ::BMKG"
    )

# ================= MAIN =================

def main():
    year = int(input("Tahun (YYYY): ").strip())
    month = int(input("Bulan (1-12): ").strip())
    yyyymm = f"{year}{month:02d}"

    print("🔎 Mengambil Menit Kirim tercepat dari JSON Node-RED ...")
    menit_map = collect_min_kirim(REGION_CODE, yyyymm)
    print(f"   → Total event dari Node-RED: {len(menit_map)}")

    event_ids = list(menit_map.keys())
    print("🔎 Mengambil parameter final dari FDSN (paralel)...")

    rows: List[Dict[str, Any]] = []

    def worker(ev_id: str) -> Optional[Dict[str, Any]]:
        xml = fetch_fdsn_final(ev_id)
        if not xml:
            return None
        q = parse_fdsn_final(xml)
        q["event_id"] = ev_id
        q["Menit Kirim"] = menit_map.get(ev_id)
        return q

    futures = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        for ev in event_ids:
            futures.append(ex.submit(worker, ev))
        for fut in as_completed(futures):
            r = fut.result()
            if r:
                rows.append(r)

    if not rows:
        print("⚠️ Tidak ada data final FDSN yang berhasil diambil.")
        return

    df = pd.DataFrame(rows)

    # Build Info Gempa
    df["Info Gempa"] = df.apply(build_info, axis=1)
    df["Menit Kirim"] = df["Menit Kirim"].astype(float).round(2)
    df["RMS"] = df["RMS"].astype(float).round(2)
    df["Gap"] = df["Gap"].astype(float).round(2)
    df["ERH"] = df["ERH"].astype(float).round(2)
    df["ERZ"] = df["ERZ"].astype(float).round(2)
    df["Mdist"] = df["Mdist"].astype(float).round(1)

    # Rapikan kolom output
    out_df = pd.DataFrame({
        "event_id": df["event_id"],
        "Info Gempa": df["Info Gempa"],
        "Menit Kirim": df["Menit Kirim"],      # sudah dalam MENIT
        "RMS (StdErr)": df["RMS"],
        "Az Gap": df["Gap"],
        "Jml Fase": df["Phase"],
        "ERH (km)": df["ERH"],
        "ERZ (km)": df["ERZ"],
        "M.dist (deg)": df["Mdist"],
    })

    outfile = f"rekap_{REGION_CODE}_{yyyymm}.xlsx"
    out_df.to_excel(outfile, index=False)

    print("\n✅ Rekap selesai.")
    print(out_df.head(10))
    print(f"\n📂 Tersimpan: {outfile}")

if __name__ == "__main__":
    main()