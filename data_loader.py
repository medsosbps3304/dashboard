# -*- coding: utf-8 -*-
"""
data_loader.py — pemuat & pengolah data XLSX / Google Spreadsheet.

Format sumber yang didukung (long format, header opsional):
    [Wilayah] [Indikator*] [Tahun] [Nilai]   (*kolom teks opsional apa pun,
                                              mis. "Kemiskinan", "IPM", dsb.)
- Satu workbook boleh berisi banyak sheet; tiap sheet = satu topik indikator.
- Nama sheet dipakai sebagai nama indikator bila kolom indikator tidak ada.
- Dimensi tambahan (mis. "Jenis Kelamin") dilebur ke nama indikator.
- Format angka Indonesia didukung: '32.917' (ribuan), '52,38' (desimal koma),
  '437.8' (desimal titik), '155249', '-', sel kosong.
"""

from __future__ import annotations

import io
import re

import pandas as pd
import requests

MISSING_TOKENS = {"", "-", "--", "---", "na", "n/a", "nil", "null", "…",
                  "...", "?", "nan"}

HEADER_TOKENS = ("wilayah", "kabupaten", "kota", "region", "daerah",
                 "indikator", "indicator", "tahun", "year", "nilai", "value",
                 "jumlah", "angka", "indeks", "data")
REGION_HINTS = ("wilayah", "kabupaten", "kota", "region", "daerah")
YEAR_HINTS = ("tahun", "year")
INDICATOR_HINTS = ("indikator", "kemiskinan", "uraian", "variabel",
                   "parameter", "jenis_data")
VALUE_HINTS = ("nilai", "value", "angka", "jumlah", "indeks", "ipm", "ipg",
               "p1", "p2", "gini", "rasio", "laju", "persen", "kontribusi",
               "inflasi")


# ---------------------------------------------------------------------------
# Util dasar
# ---------------------------------------------------------------------------
def _clean_text(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def parse_value(v):
    """Konversi nilai ke float dengan konvensi angka Indonesia."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = _clean_text(v).replace("\u00a0", "").replace(" ", "")
    if s.lower() in MISSING_TOKENS:
        return None
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    if "," in s:                                   # '1.234,56' -> 1234.56
        s = s.replace(".", "").replace(",", ".")
    elif "." in s:
        parts = s.split(".")
        # titik = pemisah ribuan, kecuali pola jelas desimal ('437.8', '71.09')
        if not (len(parts) == 2 and len(parts[1]) != 3 and len(parts[0]) <= 3):
            s = s.replace(".", "")
    try:
        val = float(s)
    except ValueError:
        return None
    if val != val:                                      # NaN ("nan", dsb.)
        return None
    return -val if neg else val


def parse_year(v):
    y = parse_value(v)
    if y is None:
        return None
    yi = int(round(y))
    return yi if 1900 <= yi <= 2100 else None


# ---------------------------------------------------------------------------
# Deteksi struktur sheet
# ---------------------------------------------------------------------------
def _detect_header_row(raw: pd.DataFrame) -> int | None:
    for i in range(min(6, len(raw))):
        cells = [_clean_text(c).lower() for c in raw.iloc[i].tolist()]
        hits = sum(any(h == c or (h and h in c) for h in HEADER_TOKENS) for c in cells)
        if hits >= 2:
            return i
    return None


def _is_year_column(col: pd.Series) -> bool:
    vals = col.dropna().head(50)
    if len(vals) < 2:
        return False
    ok = sum(parse_year(v) is not None for v in vals)
    return ok / len(vals) >= 0.6


def _looks_like_text(col: pd.Series) -> bool:
    s = col.dropna().astype(str).head(30)
    s = [x for x in s if _clean_text(x)]
    if not s:
        return False
    non_year = [x for x in s if parse_year(x) is None]
    if not non_year:
        return False
    alpha_ratio = sum(bool(re.search(r"[A-Za-zÀ-ÿ]", x)) for x in non_year) / len(non_year)
    return alpha_ratio >= 0.5


def _numeric_rate(col: pd.Series) -> float:
    vals = [v for v in col.dropna().head(60).tolist() if _clean_text(v)]
    if not vals:
        return 0.0
    return sum(parse_value(v) is not None for v in vals) / len(vals)


def _avg_text_len(col: pd.Series) -> float:
    s = [_clean_text(x) for x in col.dropna().head(40).tolist() if _clean_text(x)]
    return sum(len(x) for x in s) / len(s) if s else 0.0


def _parse_sheet(raw: pd.DataFrame, sheet_name: str) -> pd.DataFrame | None:
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.empty or raw.shape[1] < 2:
        return None
    # normalisasi label baris & kolom menjadi posisi berurutan
    raw = raw.reset_index(drop=True)
    raw.columns = range(raw.shape[1])

    header_row = _detect_header_row(raw)
    if header_row is not None:
        headers = [_clean_text(c).lower() for c in raw.iloc[header_row].tolist()]
        body = raw.iloc[header_row + 1:].reset_index(drop=True)
    else:
        headers = [""] * raw.shape[1]
        body = raw.reset_index(drop=True)
    if body.empty:
        return None

    ncols = body.shape[1]

    def header_of(j: int) -> str:
        return headers[j] if j < len(headers) else ""

    # --- kolom tahun ---
    year_col = next((j for j in range(ncols)
                     if any(h in header_of(j) for h in YEAR_HINTS)
                     and _is_year_column(body[j])), None)
    if year_col is None:
        year_col = next((j for j in range(ncols) if _is_year_column(body[j])), None)
    if year_col is None:
        return None

    text_cols = [j for j in range(ncols)
                 if j != year_col and _looks_like_text(body[j])]

    # --- kolom wilayah ---
    region_col = next((j for j in range(ncols)
                       if any(h in header_of(j) for h in REGION_HINTS)), None)
    if region_col is None:
        region_col = next((j for j in text_cols if j < year_col), None)

    # --- kolom indikator (opsional) ---
    ind_col = next((j for j in text_cols
                    if j != region_col
                    and any(h in header_of(j) for h in INDICATOR_HINTS)), None)
    if ind_col is None:
        scored = [(round(_avg_text_len(body[j]), 1), j) for j in text_cols
                  if j != region_col]
        if scored:
            best_len, best_j = max(scored)
            ind_col = best_j if best_len >= 14 else None

    # --- dimensi aux (mis. Jenis Kelamin): teks lain yang bervariasi ---
    aux_cols = []
    for j in text_cols:
        if j in {region_col, ind_col}:
            continue
        uniq = {_clean_text(x) for x in body[j].dropna().tolist() if _clean_text(x)}
        if len(uniq) > 1:
            aux_cols.append(j)

    # --- kolom nilai: skor keterisian angka + bonus kata-kunci header ---
    reserved = {year_col, region_col, ind_col, *aux_cols}
    candidates = [j for j in range(ncols) if j not in reserved]

    def value_score(j: int):
        rate = _numeric_rate(body[j])
        bonus = 0.05 if any(h in header_of(j) for h in VALUE_HINTS) else 0.0
        return min(rate + bonus, 1.0)

    value_col = max(candidates, key=lambda j: (value_score(j), j)) \
        if candidates else None
    if value_col is None or _numeric_rate(body[value_col]) < 0.5:
        return None

    # --- aux yang benar-benar bervariasi saja sudah difilter di atas ---

    rows = []
    for _, r in body.iterrows():
        tahun = parse_year(r[year_col])
        if tahun is None:
            continue
        w_raw = _clean_text(r[region_col]) if region_col is not None else ""
        wilayah = w_raw.title() or _clean_text(sheet_name).title() or "Wilayah"
        ind_base = _clean_text(r[ind_col]) if ind_col is not None else ""
        if not ind_base:
            ind_base = _clean_text(sheet_name) or "Indikator"
        aux_parts = [_clean_text(r[j]) for j in aux_cols if _clean_text(r[j])]
        indikator = ind_base + (f" ({' · '.join(aux_parts)})" if aux_parts else "")
        rows.append({
            "wilayah": wilayah,
            "indikator": indikator,
            "tahun": tahun,
            "nilai": parse_value(r[value_col]),
        })

    df = pd.DataFrame(rows)
    return df if not df.empty else None


# ---------------------------------------------------------------------------
# API publik
# ---------------------------------------------------------------------------
def load_workbook_df(fileobj) -> pd.DataFrame:
    """Baca seluruh sheet workbook menjadi tabel rapi."""
    frames = []
    xls = pd.ExcelFile(fileobj)
    for name in xls.sheet_names:
        try:
            raw = pd.read_excel(xls, sheet_name=name, header=None)
        except Exception:                               # noqa: BLE001
            continue
        parsed = _parse_sheet(raw, name)
        if parsed is not None:
            frames.append(parsed)
    if not frames:
        raise ValueError(
            "Tidak ada data terbaca. Pastikan format kolom: "
            "[Wilayah | Indikator | Tahun | Nilai].")
    df = pd.concat(frames, ignore_index=True).drop_duplicates()
    return df.sort_values(["wilayah", "indikator", "tahun"]).reset_index(drop=True)


def load_folder(folder: str, allowed: list[str] | None = None,
                ) -> tuple[pd.DataFrame, list[str]]:
    """Baca file .xlsx/.xls dalam sebuah folder (tanpa cache).

    allowed : daftar nama file; urutan = prioritas saat konflik nilai
              antar-file (baris dari file pertama yang punya nilai menang).
              None -> semua *.xlsx/*.xls di folder dibaca (perilaku lama).
    """
    from pathlib import Path
    base = Path(folder)
    if allowed is not None:
        files = [base / name for name in allowed]
    else:
        files = (sorted([*base.glob("*.xlsx"), *base.glob("*.xls")])
                 if base.exists() else [])
    errors, frames = [], []
    for f in files:
        if not f.exists() or f.name.startswith("~$"):
            continue
        try:
            with open(f, "rb") as fh:
                frames.append(load_workbook_df(fh))
        except Exception as exc:                        # noqa: BLE001
            errors.append(f"{f.name}: {exc}")
    if not frames:
        raise FileNotFoundError(
            "; ".join(errors) or f"Tidak ada file XLSX di {folder}.")
    df = pd.concat(frames, ignore_index=True)
    # 1) buang salinan identik lintas-file (mis. snapshot live == Data_BPS);
    # 2) selesaikan konflik (w+i+t sama, nilai beda): baris bernilai menang,
    #    sisanya mengikuti urutan prioritas file (sort stabil).
    df = df.drop_duplicates(subset=["wilayah", "indikator", "tahun", "nilai"])
    df["_has_val"] = df["nilai"].notna()
    df = (df.sort_values(["wilayah", "indikator", "tahun", "_has_val"],
                         ascending=[True, True, True, False], kind="stable")
            .drop_duplicates(subset=["wilayah", "indikator", "tahun"],
                             keep="first")
            .drop(columns="_has_val"))
    return df.reset_index(drop=True), errors


def fetch_google_sheet(url: str | None = None) -> pd.DataFrame:
    """Ambil workbook spreadsheet publik dan parse langsung."""
    if not url:
        raise ValueError(
            "URL spreadsheet belum diberikan. Atur SHEET_XLSX_URL (atau "
            "SHEET_ID) melalui Streamlit Secrets, atau teruskan argumen `url`.")
    resp = requests.get(url, timeout=60, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    return load_workbook_df(io.BytesIO(resp.content))


def download_xlsx_bytes(url: str | None = None) -> bytes:
    if not url:
        raise ValueError(
            "URL spreadsheet belum diberikan. Atur SHEET_XLSX_URL (atau "
            "SHEET_ID) melalui Streamlit Secrets, atau teruskan argumen `url`.")
    resp = requests.get(url, timeout=60, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    return resp.content


# ---------------------------------------------------------------------------
# Label & klasifikasi indikator
# ---------------------------------------------------------------------------
def kind_of(indikator: str) -> str:
    low = (indikator or "").lower()
    if "persentase" in low and "miskin" in low:
        return "ppm"
    if "garis" in low and "kemiskinan" in low:
        return "gk"
    return "other"


def short_label(indikator: str) -> str:
    low = (indikator or "").lower()
    if "persentase" in low and "miskin" in low:
        return "Persentase Penduduk Miskin (%)"
    if "garis" in low and "kemiskinan" in low:
        return "Garis Kemiskinan (Rp/kapita/bln)"
    return indikator


# ---------------------------------------------------------------------------
# Pemetaan 21 indikator kanonik (urutan resmi sesuai daftar indikator)
# ---------------------------------------------------------------------------
CANONICAL: list[str] = [
    "Indeks Kemahalan Konstruksi",
    "Indeks Pembangunan Manusia",
    "Tingkat Kemiskinan",
    "Inflasi Kota Purwokerto",
    "Indeks Ketimpangan Gender",
    "Jumlah Lansia",
    "Jumlah Penduduk Miskin",
    "Tingkat Pengangguran Terbuka",
    "PDRB Perkapita",
    "Indeks Pemberdayaan Gender",
    "Tingkat Kedalaman Kemiskinan",
    "Indeks Pembangunan Gender",
    "Gini Ratio",
    "Tingkat Keparahan Kemiskinan",
    "Kontribusi Pertanian Terhadap PDRB",
    "Laju Pertumbuhan Penduduk",
    "Rasio Ekspor Terhadap PDRB",
    "Pertumbuhan Ekonomi",
    "Tingkat Partisipasi Angkatan Kerja",
    "Jumlah Balita",
    "Rasio Investasi Terhadap PDRB",
]

_CANON_RULES: list[tuple[str, tuple[str, ...]]] = [
    ("Indeks Kemahalan Konstruksi",     ("kemahalan konstruksi",)),
    ("Indeks Pembangunan Manusia",      ("pembangunan manusia",)),
    ("Tingkat Kemiskinan",              ("garis kemiskinan", "persentase penduduk miskin")),
    ("Inflasi Kota Purwokerto",         ("inflasi",)),
    ("Indeks Ketimpangan Gender",       ("ketimpangan gender",)),
    ("Jumlah Lansia",                   ("lansia",)),
    ("Jumlah Penduduk Miskin",          ("jumlah penduduk miskin",)),
    ("Tingkat Pengangguran Terbuka",    ("pengangguran",)),
    ("PDRB Perkapita",                  ("pdrb perkapita", "pendapatan per kapita",
                                         "pendapatan perkapita")),
    ("Indeks Pemberdayaan Gender",      ("pemberdayaan gender",)),
    ("Tingkat Kedalaman Kemiskinan",    ("kedalaman",)),
    ("Indeks Pembangunan Gender",       ("pembangunan gender",)),
    ("Gini Ratio",                      ("gini",)),
    ("Tingkat Keparahan Kemiskinan",    ("keparahan",)),
    ("Kontribusi Pertanian Terhadap PDRB", ("kontribusi pertanian",)),
    ("Laju Pertumbuhan Penduduk",       ("pertumbuhan penduduk",)),
    ("Rasio Ekspor Terhadap PDRB",      ("rasio ekspor",)),
    ("Pertumbuhan Ekonomi",             ("pertumbuhan ekonomi",)),
    ("Tingkat Partisipasi Angkatan Kerja", ("partisipasi angkatan",)),
    ("Jumlah Balita",                   ("balita",)),
    ("Rasio Investasi Terhadap PDRB",   ("rasio investasi",)),
]


def canon_key(indikator: str) -> str | None:
    """Petakan nama indikator hasil parsing -> nama kanonik (atau None)."""
    low = (indikator or "").lower()
    for name, keywords in _CANON_RULES:
        if any(k in low for k in keywords):
            return name
    return None
