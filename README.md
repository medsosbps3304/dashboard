# Dashboard Statistik Daerah — BPS Banjarnegara

Dashboard interaktif berbasis **Streamlit** bertema **aksen resmi BPS**
(biru tua `#0A3D6E` + aksen merah `#E23A34`). Logo resmi BPS (`lambang_bps.svg`)
ditanamkan di sidebar dan banner; logo `bara-mendoan.png` dipakai sebagai
favicon sekaligus logo kedua di banner judul.

Sumber data diambil **langsung dari Google Spreadsheet** (berkas induk) setiap
sesi aplikasi dijalankan, lalu disimpan sebagai salinan lokal
`data/_live_snapshot.xlsx`. Bila jaringan tidak tersedia, dashboard menggunakan
salinan snapshot terakhir secara transparan.

## Fitur

| Fitur | Keterangan |
|---|---|
| Logo BPS & Bara Mendoan | Tertanam di sidebar & banner (SVG/PNG, tanpa dependensi eksternal saat runtime) |
| Kartu KPI fokus dinamis | Mengikuti dua indikator teratas yang dipilih + observasi terfilter + periode tersedia |
| Peta 21 indikator | Grid mini-kartu seluruh indikator terpantau; yang belum terisi tampil redup *"belum tersedia"* |
| Grafik tren | Interaktif (Plotly), garis atau batang |
| Wilayah terkunci | Data ditampilkan untuk **Kabupaten Banjarnegara** saja (pemilihan wilayah dinonaktifkan) |
| Komposisi penduduk | Diagram pie *penduduk miskin vs sejahtera* dari Persentase Penduduk Miskin; tampil **hanya saat indikator *Garis Kemiskinan* dipilih** |
| Dinamika Kemiskinan | Dual-axis: batang % penduduk miskin vs garis kemiskinan Rp; tampil saat varian *Tingkat Kemiskinan* dipilih |
| Scatter Ekonomi–Pengangguran | Sebaran tahunan Pertumbuhan Ekonomi vs Tingkat Pengangguran Terbuka dengan garis rata-rata tiap sumbu |
| Perbandingan tumpang-tindih | Dua indikator dalam satu kanvas (Garis/Area/Batang); satuan berbeda otomatis ke sumbu kanan |
| Tabel rinci | Format angka Indonesia + unduh CSV |
| Sinkron otomatis | Tiap sesi dimulai, data terbaru diambil dari spreadsheet induk; bila jaringan gagal dipakai salinan lokal `data/` |

Antarmuka tidak menampilkan elemen konektor maupun keterangan sumber data.

## Cara Menjalankan

```bash
pip install -r requirements.txt   # sekali saja
streamlit run app.py
```

Browser terbuka otomatis di `http://localhost:8501`.

## Sumber & Pembaruan Data

1. Tambah/ubah baris tahun baru pada **Google Spreadsheet induk** (terhubung
   lewat `SHEET_XLSX_URL` di `data_loader.py`).
2. Buka (atau muat ulang) dashboard — data terbaru **otomatis** diambil dan
   disalin ke `data/_live_snapshot.xlsx`.
3. Tanpa internet? Dashboard memakai salinan lokal terakhir secara transparan.

Tidak ada tombol atau langkah sinkronisasi manual; tidak ada kode yang perlu
diubah untuk menambah periode baru. Penggabungan beberapa berkas Excel lokal
tidak lagi digunakan — seluruh data berasal dari spreadsheet induk.

## Konfigurasi

ID spreadsheet **wajib** diisi melalui **Streamlit Secrets** (dibaca di `app.py`
saat startup) — tidak ada nilai bawaan yang ditanam di source code. Secara lokal
letakkan di `.streamlit/secrets.toml` (sudah di-gitignore); di Streamlit Cloud,
isikan lewat panel Secrets. Bila `SHEET_XLSX_URL` tidak diisi, URL ekspor akan
dibangun otomatis dari `SHEET_ID`.

```toml
# .streamlit/secrets.toml
SHEET_ID = "ISI_DENGAN_ID_SPREADSHEET_ANDA"
# SHEET_XLSX_URL bersifat opsional; bila ada akan dipakai langsung
# SHEET_XLSX_URL = "https://docs.google.com/spreadsheets/d/ISI_ID/export?format=xlsx"
```

Tanpa `SHEET_ID` di secrets, aplikasi berhenti saat startup dengan pesan error
yang jelas.

## Struktur Proyek

```
Dashboard-BPSBNA-main/
├─ app.py                  # aplikasi Streamlit (UI) + CSS tema aksen BPS
├─ data_loader.py          # parser XLSX / Google Spreadsheet + pemetaan 21 indikator kanonik
├─ requirements.txt        # dependensi Python
├─ README.md
├─ assets/
│  ├─ lambang_bps.svg      # logo resmi BPS (Wikimedia Commons)
│  └─ bara-mendoan.png     # logo kedua (favicon + banner, klik = muat ulang)
└─ data/
   └─ _live_snapshot.xlsx  # salinan snapshot sinkronisasi live (dibuat otomatis)
```

### Dependensi (`requirements.txt`)

`streamlit>=1.36`, `pandas>=2.0`, `plotly>=5.20`, `openpyxl>=3.1`, `requests>=2.31`

## Format Data yang Didukung

Long format per baris (header opsional, nama kolom nilai bebas):

| Wilayah      | Indikator                          | Tahun | Nilai    |
|--------------|------------------------------------|-------|----------|
| Banjarnegara | Garis Kemiskinan (Rp/kapita/bln)   | 1996  | 32.917   |
| Banjarnegara | Persentase Penduduk Miskin (persen)| 1999  | 52,38    |

Aturan pembacaan otomatis:

- Satu sheet = satu topik indikator; nama sheet dipakai bila kolom indikator
  tidak ada.
- Dimensi tambahan (**Jenis Kelamin**) dilebur ke nama indikator →
  *Indeks Pembangunan Manusia (Laki-laki)*.
- Angka gaya Indonesia: `32.917` = 32917 · `52,38` = 52.38 · `437.8` = 437.8.
- Sel kosong / `-` menjadi *missing* (grafik putus, tidak error).
- Sheet kosong & baris sisa diabaikan.
- Baris identik antar-pengambilan dibuang agar angka tidak terhitung dobel.

### Peta 21 Indikator Kanonik

Urutan kartu mengikuti daftar indikator resmi (lihat `CANONICAL` di
`data_loader.py`). Parser memetakan hasil bacaan sheet ke nama kanonik lewat
`canon_key()`. Beberapa indikator memiliki dua mode kartu:

- **IPM L/P** tampil berdampingan dalam satu kartu (*Laki-laki* kiri,
  *Perempuan* kanan).
- Kartu *Tingkat Kemiskinan* memakai **toggle ⇄ 2 varian** antara Garis
  Kemiskinan dan % Penduduk Miskin.
- Kartu KPI *Periode Tersedia* menyesuaikan rentang tahun indikator yang dipilih.

> Catatan: Gauge kategori IPM saat ini **tidak aktif** (fungsinya dikomentari di
> `app.py`); indikator IPM tetap muncul pada grid 21 indikator.

## Palet Warna

| Peran | Warna |
|---|---|
| Biru utama BPS | `#0A3D6E` |
| Biru sekunder | `#1976C5` |
| Aksen merah logo | `#E23A34` |
| Hijau (turun = baik) | `#00FF66` |
| Amber (garis rata-rata) | `#F2A93B` |
| Tinta teks | `#12263F` |
| Abu keabu-abuan (teks sekunder) | `#5B7089` |
| Latar kabut | `#EAF3FB` |
| Garis pembatas | `#E3EBF4` |

## Pemecahan Masalah

- **"Belum ada data yang dapat dimuat"** → buka dengan koneksi internet
  tersedia agar salinan lokal pertama berhasil dibuat (diunduh dari spreadsheet
  induk ke `data/_live_snapshot.xlsx`).
- **Angka aneh** → parser mengikuti konvensi Indonesia (titik ribuan,
  koma desimal); pastikan sel angka tidak bercampur teks lain.
- **Logo hilang** → pastikan `assets/lambang_bps.svg` dan `assets/bara-mendoan.png`
  berada di folder `assets/` relatif terhadap `app.py`.
