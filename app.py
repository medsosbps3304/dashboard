# -*- coding: utf-8 -*-
"""
Dashboard Statistik Daerah — BPS Banjarnegara
=============================================
Dashboard Streamlit bertema aksen resmi BPS.

- Data dimuat OTOMATIS setiap sesi aplikasi dimulai dan diambil HANYA dari
  spreadsheet induk (live), lalu disimpan sebagai salinan lokal. Bila jaringan
  tidak tersedia, dipakai salinan snapshot terakhir di folder `data/`.
- Tidak ada elemen konektor/sumber data yang tampil di antarmuka.

Menjalankan:  streamlit run app.py
"""

from __future__ import annotations

import base64
import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from data_loader import (CANONICAL, canon_key, download_xlsx_bytes, kind_of,
                         load_workbook_df, short_label)


# ---------------------------------------------------------------------------
# Rahasia: dibacakan dari Streamlit Secrets (lokal: .streamlit/secrets.toml;
# Streamlit Cloud: panel Secrets). Tidak ada nilai bawaan di source code.
# ---------------------------------------------------------------------------
def _secret(key: str, default: str | None = None) -> str:
    try:
        val = st.secrets[key]
    except Exception:
        val = None
    if not val:
        if default is None:
            raise RuntimeError(
                f"Rahasia '{key}' tidak ditemukan di Streamlit Secrets. "
                f"Isikan di .streamlit/secrets.toml (lokal) atau panel Secrets "
                f"(Streamlit Cloud).")
        return default
    return val


_SHEET_ID = _secret("SHEET_ID")            # wajib
SHEET_XLSX_URL = _secret(
    "SHEET_XLSX_URL",
    f"https://docs.google.com/spreadsheets/d/{_SHEET_ID}/export?format=xlsx",
)


# ---------------------------------------------------------------------------
# Konfigurasi halaman & palet BPS
# ---------------------------------------------------------------------------
# Gambar "bara mendoan" (PNG transparan): favicon tab browser sekaligus
# logo kedua di banner judul. Fallback emoji bila file tidak ada.
MENDOAN_PATH = Path(__file__).parent / "assets" / "bara-mendoan.png"

st.set_page_config(
    page_title="NewBaramendoan",
    page_icon=str(MENDOAN_PATH) if MENDOAN_PATH.exists() else "📊",
    layout="wide",
    initial_sidebar_state="collapsed",   # sidebar tersembunyi saat dibuka
)

C = {
    "navy":  "#0A3D6E",   # biru tua resmi BPS
    "blue":  "#1976C5",
    "red":   "#E23A34",   # aksen merah logo BPS
    "green": "#00FF66",
    "amber": "#F2A93B",
    "ink":   "#12263F",
    "muted": "#5B7089",
    "mist":  "#EAF3FB",
    "line":  "#E3EBF4",
}
SERIES_PALETTE = [C["navy"], C["red"], C["green"], C["amber"],
                  C["blue"], "#7E57C2", "#00897B", "#D81B60"]
FOCUS_BAR_COLORS = [C["red"], C["navy"], C["green"], C["blue"]]

DATA_DIR = Path(__file__).parent / "data"
# Snapshot hasil sinkronisasi live; nama lama "kemiskinan_terbaru.xlsx" menyesatkan
# karena isinya seluruh indikator, bukan hanya kemiskinan.
SYNC_TARGET = DATA_DIR / "_live_snapshot.xlsx"
# Data kini hanya dari spreadsheet; prioritas antar-file lokal tidak dipakai lagi.
# FILE_PRIORITY = [SYNC_TARGET.name, "Data_BPS.xlsx", "bps_api_tambahan.xlsx"]
LOGO_PATH = Path(__file__).parent / "assets" / "lambang_bps.svg"

# ===========================================================================
# Util tampilan
# ===========================================================================
def fmt_num(x, dec: int = 2) -> str:
    """Format angka gaya Indonesia: ribuan '.', desimal ','."""
    if x is None or pd.isna(x):
        return "-"
    s = f"{x:,.{dec}f}"
    return s.replace(",", "@").replace(".", ",").replace("@", ".")


def fmt_rp(x) -> str:
    return "Rp " + fmt_num(x, 0)


def smart_dec(x) -> int:
    if x is None or pd.isna(x):
        return 0
    return 0 if abs(x) >= 1000 else 2


def is_percent_family(indikator: str) -> bool:
    return "(persen)" in indikator.lower() or "(%)" in indikator.lower()


def delta_display(v: float, vp: float | None, indikator: str):
    """Kembalikan (delta_tampil, satuan_chip) atau (None, '')."""
    if vp is None or pd.isna(vp) or pd.isna(v):
        return None, ""
    d = v - vp
    if is_percent_family(indikator):
        return d, " poin"
    if vp != 0:
        return d / abs(vp) * 100, "%"
    return d, ""


def delta_chip(delta, unit="", dec=2) -> str:
    if delta is None or pd.isna(delta):
        return '<span class="chip chip-muted">data awal</span>'
    arrow = "▼" if delta < 0 else ("▲" if delta > 0 else "▬")
    cls = "chip-down" if delta < 0 else ("chip-up" if delta > 0 else "chip-flat")
    sign = "" if delta < 0 else "+"
    return (f'<span class="chip {cls}">{arrow} {sign}{fmt_num(delta, dec)}'
            f'{unit}</span>')


def bps_logo(size: int = 40) -> str:
    """Logo BPS tertanam (data-URI) dengan fallback emoji."""
    if LOGO_PATH.exists():
        b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode()
        return (f'<img src="data:image/svg+xml;base64,{b64}" width="{size}" '
                f'height="{size}" alt="Logo BPS" '
                f'style="display:block;border-radius:10px"/>')
    return f'<div style="font-size:{size * 0.7:.0f}px">📊</div>'


def mendoan_logo(size: int = 56) -> str:
    """Gambar bara-mendoan tertanam (data-URI) dengan fallback emoji."""
    if MENDOAN_PATH.exists():
        b64 = base64.b64encode(MENDOAN_PATH.read_bytes()).decode()
        return (f'<img src="data:image/png;base64,{b64}" width="{size}" '
                f'height="{size}" alt="Bara Mendoan" '
                f'style="display:block;border-radius:10px"/>')
    return f'<div style="font-size:{size * 0.7:.0f}px">🔥</div>'


def data_updated_at() -> datetime:
    mtimes = [f.stat().st_mtime for f in DATA_DIR.glob("*.xls*") if f.exists()]
    return datetime.fromtimestamp(max(mtimes)) if mtimes else datetime.now()


# ===========================================================================
# CSS — tema aksen BPS
# ===========================================================================
st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

:root {{
  --navy:{C['navy']}; --blue:{C['blue']}; --red:{C['red']};
  --green:{C['green']}; --ink:{C['ink']}; --muted:{C['muted']};
  --mist:{C['mist']}; --line:{C['line']};
}}
html, body, [class*="css"], .stApp {{
  font-family:'Plus Jakarta Sans','Segoe UI',system-ui,sans-serif;
  color:var(--ink);
}}
.stApp {{ background:linear-gradient(180deg,#F4F8FC 0%,#F7FAFD 100%); }}

#MainMenu {{ visibility:hidden; }}
footer {{ visibility:hidden; }}
header[data-testid="stHeader"] {{ background:transparent; }}
.block-container {{ padding-top:1.4rem; padding-bottom:3rem; max-width:1240px; }}

/* ---------- banner judul ---------- */
.hero {{
  background:linear-gradient(115deg,var(--navy) 0%,#10528F 55%,var(--blue) 100%);
  border-radius:20px; padding:22px 30px 24px; color:#fff;
  box-shadow:0 12px 30px rgba(10,61,110,.25); margin-bottom:22px;
  position:relative; overflow:hidden;
  display:flex; flex-direction:column; align-items:flex-start; gap:14px;
}}
.hero::after {{
  content:''; position:absolute; right:-60px; top:-70px; width:260px; height:260px;
  border-radius:50%; background:rgba(255,255,255,.07);
}}
/* baris logo: bara-mendoan + BPS berdampingan, di kiri atas banner */
.hero .hero-logos {{
  display:flex; align-items:center; justify-content:flex-start; gap:16px; z-index:1;
}}
/* logo tanpa wadah — container transparan, gambar tampil apa adanya
   di atas gradasi banner */
.hero .hero-logo-plain {{
  flex:0 0 auto; display:flex; align-items:center; justify-content:center;
  background:transparent; z-index:1;
}}
.hero .hero-logo-plain img {{
  display:block; filter:drop-shadow(0 4px 10px rgba(0,0,0,.35));
}}
/* logo dapat diklik: tanpa garis bawah, sedikit membesar saat hover */
.hero a.hero-logo-plain {{ text-decoration:none; cursor:pointer; }}
.hero a.hero-logo-plain:hover {{ transform:scale(1.06); }}
.hero a.hero-logo-plain {{ transition:transform .15s ease; }}
/* teks di kiri, stempel tanggal + badge di kanan */
.hero .hero-main {{
  display:flex; align-items:center; gap:20px; width:100%; z-index:1;
}}
.hero .hero-text {{ flex:1 1 auto; min-width:0; }}
.hero .hero-side {{
  flex:0 0 auto; display:flex; flex-direction:column; align-items:flex-end; gap:8px;
}}
.hero h1 {{ font-size:1.45rem; font-weight:800; margin:0 0 4px; }}
.hero p {{ margin:0; opacity:.88; font-size:.86rem; }}
/* subjudul tagline: melekat rapat di bawah judul, beda level dari baris konten */
.hero p.subjudul {{
  font-style:italic; font-weight:600; opacity:.95;
  font-size:.8rem; letter-spacing:.015em; margin:0 0 10px;
}}
.hero p.fokus {{ margin-top:4px; opacity:.75; font-size:.78rem; }}
.hero .badge-stack {{
  display:flex; flex-direction:column; align-items:flex-end; gap:6px;
}}
.hero .badge-stack span {{
  display:inline-flex; align-items:center;
  background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.25);
  padding:5px 12px; border-radius:999px; font-size:.73rem; font-weight:600;
}}
.hero .hero-stamp {{
  flex:0 0 auto; z-index:1; white-space:nowrap;
  background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.25);
  padding:6px 14px; border-radius:999px; font-size:.73rem; font-weight:600;
}}

/* ---------- hero di layar sempit (mobile) ----------
   Deretan logo di kiri-atas, teks di bawahnya (rata kiri),
   stempel tanggal + badge mengalir di bawah teks. */
@media (max-width: 768px) {{
  .hero {{
    padding:16px 18px 18px; gap:12px;
  }}
  .hero .hero-logos {{ gap:12px; }}
  .hero .hero-logo-plain img {{ width:60px; height:60px; }}
  .hero .hero-main {{ flex-direction:column; align-items:flex-start; gap:8px; }}
  .hero .hero-side {{ align-items:flex-start; }}
  .hero h1 {{ font-size:1.12rem; }}
  .hero p {{ display:inline; font-size:.8rem; }}
  .hero p.subjudul {{ font-size:.78rem; margin:0 10px 0 6px; }}
  .hero p.fokus {{ font-size:.76rem; }}
  .hero p.fokus::before {{ content:' · '; opacity:.7; }}
  .hero .badge-stack {{ align-items:flex-start; gap:5px; }}
  /* footer: kolom kontak & sosial ditumpuk di layar sempit */
  .site-footer .foot-cols {{ flex-direction:column; gap:18px; }}
  .site-footer .foot-col {{ min-width:0; }}
}}

/* ---------- kartu KPI ---------- */
.kpi-grid {{ display:flex; gap:14px; flex-wrap:wrap; margin-bottom:6px; }}
.kpi {{
  flex:1 1 220px; min-width:210px; background:#fff; border-radius:16px;
  border:1px solid var(--line); box-shadow:0 4px 16px rgba(16,42,77,.06);
  padding:18px 20px 16px; position:relative; overflow:hidden;
}}
.kpi::before {{
  content:''; position:absolute; left:0; top:0; bottom:0; width:4px;
  background:var(--bar,var(--navy));
}}
.kpi .lbl {{
  font-size:.66rem; font-weight:700; letter-spacing:.1em;
  text-transform:uppercase; color:var(--muted); margin-bottom:6px;
}}
.kpi .val {{ font-size:1.75rem; font-weight:800; line-height:1.15; }}
.kpi .sub {{ margin-top:8px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }}
.kpi .hint {{ font-size:.72rem; color:var(--muted); }}

/* ---------- grid mini-kartu 21 indikator ---------- */
.mini-grid {{
  display:grid; grid-template-columns:repeat(auto-fill,minmax(196px,1fr));
  gap:12px; margin-bottom:8px;
}}
.mini {{
  background:#fff; border:1px solid var(--line); border-radius:14px;
  padding:13px 15px 12px; box-shadow:0 3px 10px rgba(16,42,77,.05);
  border-left:4px solid var(--bar,var(--navy));
}}
.mini .lbl {{
  font-size:.62rem; font-weight:700; letter-spacing:.07em;
  text-transform:uppercase; color:var(--muted); line-height:1.35;
  min-height:30px; display:flex; align-items:flex-end; margin-bottom:6px;
}}
.mini .val {{ font-size:1.22rem; font-weight:800; line-height:1.2;
             word-break:break-word; }}
.mini .sub {{ margin-top:6px; display:flex; gap:6px; align-items:center;
             flex-wrap:wrap; font-size:.68rem; color:var(--muted); }}
.mini.empty {{ opacity:.62; background:#FBFCFE; border-style:dashed;
              border-left-style:solid; }}
.mini.empty .val {{ color:var(--muted); font-size:.85rem; font-weight:600; }}
/* --- dua varian berdampingan dalam satu kartu (mis. IPM L/P) --- */
.duo {{ display:flex; gap:8px; margin-top:2px; }}
.duo-item {{ flex:1 1 0; min-width:0; }}
.duo-item .val {{ font-size:1.05rem; word-break:keep-all; }}
.duo-lbl {{
  font-size:.64rem; font-weight:600; color:var(--muted); margin-top:2px;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
}}
/* --- toggle varian di dalam kartu (checkbox + CSS, tanpa JS) --- */
.mini input.vswitch {{ display:none; }}
.mini .face {{ display:none; }}
.mini .face.f0 {{ display:block; }}
.mini input.vswitch:checked ~ .face.f0 {{ display:none; }}
.mini input.vswitch:checked ~ .face.f1 {{ display:block; }}
.mini .vtoggle {{
  cursor:pointer; user-select:none;
  display:inline-flex; align-items:center; gap:5px;
  margin-top:8px; padding:3px 11px; border-radius:999px;
  font-size:.66rem; font-weight:700; color:var(--navy);
  background:var(--mist); border:1px solid var(--line);
  transition:background .15s ease;
}}
.mini .vtoggle:hover {{ background:#dcebfa; }}

.chip {{
  display:inline-block; padding:3px 10px; border-radius:999px;
  font-size:.71rem; font-weight:700;
}}
.chip-down {{ background:#E5F5EE; color:var(--green); }}
.chip-up   {{ background:#FDEAE9; color:var(--red); }}
.chip-flat {{ background:#EDF1F6; color:var(--muted); }}
.chip-muted{{ background:#EDF1F6; color:var(--muted); }}
.chip-info {{ background:var(--mist); color:var(--navy); }}
.chip-year {{ background:var(--mist); color:var(--navy); padding:2px 8px;
             font-size:.66rem; }}

/* ---------- judul seksi & kartu grafik ---------- */
.section-title {{ display:flex; align-items:center; gap:10px; margin:26px 0 12px; }}
.section-title .dot {{
  width:10px; height:10px; border-radius:3px; background:var(--red);
  transform:rotate(45deg); flex:0 0 auto;
}}
.section-title h3 {{ margin:0; font-size:1.02rem; font-weight:800; }}
.section-title small {{
  color:var(--muted); font-weight:500; margin-left:auto; font-size:.76rem;
}}

/* ---------- sidebar ---------- */
section[data-testid="stSidebar"] {{
  background:linear-gradient(180deg,#0A3D6E 0%,#0D476F 100%);
}}
section[data-testid="stSidebar"] * {{ color:#FFFFFF !important; }}
/* tombol pelipat sidebar dinamis (Streamlit >=1.50):
   - sidebar tertutup -> ">>" (stExpandSidebarButton) biru merek
   - sidebar terbuka  -> "<<" (stSidebarCollapseButton) putih di atas navy
   Kedua elemen hanya tampil sesuai keadaan sidebar, jadi warna ikut berubah. */
button[data-testid="baseButtonHeader"],
[data-testid="stExpandSidebarButton"],
[data-testid="stExpandSidebarButton"] * {{
  color:#1976C5 !important;
}}
button[data-testid="baseButtonHeader"] svg,
[data-testid="stExpandSidebarButton"] svg {{
  fill:#1976C5 !important; stroke:#1976C5 !important;
}}
section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"],
section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] * {{
  color:#FFFFFF !important;
}}
section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] svg {{
  fill:#FFFFFF !important; stroke:#FFFFFF !important;
}}
section[data-testid="stSidebar"] hr {{ border-color:rgba(255,255,255,.15); }}
.sidebar-brand {{ display:flex; align-items:center; gap:12px; padding:6px 4px 14px; }}
.sidebar-brand .logo {{
  width:52px; height:52px; border-radius:14px; background:#fff;
  display:flex; align-items:center; justify-content:center;
  box-shadow:0 6px 14px rgba(0,0,0,.25);
}}
.sidebar-brand b {{ font-size:1rem; line-height:1.25; display:block; }}
.sidebar-brand span {{ font-size:.7rem; opacity:.75; }}
.side-caption {{ font-size:.73rem !important; opacity:.65; line-height:1.55; }}

div[data-baseweb="select"] > div {{ border-radius:10px !important; }}
[data-testid="stDataFrame"] {{ border-radius:14px; overflow:hidden; }}
.stDownloadButton > button {{
  border-radius:10px; font-weight:700; width:100%;
  background:var(--navy); color:#fff; border:none;
}}
/* ---------- footer media sosial: senada header, full-bleed ujung ke ujung ---------- */
.stApp {{ overflow-x:hidden; }}   /* cegah scroll horizontal akibat lebar 100vw */
.site-footer {{
  width:100vw;
  margin:34px calc(50% - 50vw) -3rem;  /* tembus container kiri-kanan + tepi bawah */
  border-radius:0;                     /* menyentuh tepi -> tanpa sudut bulat */
  padding:22px max(24px, calc(50vw - 596px)) 18px;
  background:linear-gradient(115deg,var(--navy) 0%,#10528F 55%,var(--blue) 100%);
  color:#FFFFFF;
  box-shadow:0 -8px 24px rgba(10,61,110,.12);
  position:relative; overflow:hidden;
}}
.site-footer .foot-title {{
  font-size:.72rem; font-weight:800; letter-spacing:.12em;
  text-transform:uppercase; opacity:.85; margin:0 0 10px;
}}
.site-footer .soc-grid {{ display:flex; flex-wrap:wrap; gap:10px; }}
.site-footer a.soc {{
  display:inline-flex; align-items:center; gap:7px;
  background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.25);
  padding:7px 14px; border-radius:999px; font-size:.78rem; font-weight:700;
  color:#FFFFFF; text-decoration:none; transition:background .15s ease;
}}
.site-footer a.soc:hover {{ background:rgba(255,255,255,.3); }}
.site-footer a.soc svg {{ flex:0 0 auto; }}
.site-footer .footnote-line {{
  margin-top:14px; font-size:.74rem; opacity:.72; text-align:center;
}}
/* --- susunan dua kolom: media sosial + kontak --- */
.site-footer .foot-cols {{
  display:flex; gap:34px; flex-wrap:wrap; align-items:flex-start;
}}
.site-footer .foot-col {{ flex:1 1 300px; min-width:260px; }}
.site-footer .contact {{
  font-size:.78rem; line-height:1.9; opacity:.95;
}}
.site-footer .contact b {{ font-size:.82rem; }}
.site-footer .contact a {{ color:#FFFFFF; font-weight:700; }}

/* ===== Slider Rentang Tahun: kartu gelap transparan, teks putih ===== */
section[data-testid="stSidebar"] div[data-testid="stSlider"],
section[data-testid="stSidebar"] div.stSlider {{
  background:rgba(255,255,255,.08);
  border:1px solid rgba(255,255,255,.16);
  border-radius:14px;
  padding:14px 18px 34px;
  margin:4px 0 16px;
}}
/* SEMUA teks di dalam kartu slider -> putih (judul, nilai, gelembung) */
section[data-testid="stSidebar"] div[data-testid="stSlider"] *,
section[data-testid="stSidebar"] div.stSlider * {{
  color:#FFFFFF !important;
}}
/* pegangan slider putih agar terlihat di atas latar navy */
section[data-testid="stSidebar"] div[data-testid="stSlider"] [role="slider"],
section[data-testid="stSidebar"] div.stSlider [role="slider"] {{
  background-color:#FFFFFF;
  border:none;
}}

/* ===== Widget pilihan (Wilayah / Indikator): gelap + teks putih ===== */
section[data-testid="stSidebar"] div[data-baseweb="select"],
section[data-testid="stSidebar"] div[data-baseweb="select"] > div {{
  background-color:rgba(255,255,255,.08);
  border-color:rgba(255,255,255,.28);
}}
section[data-testid="stSidebar"] div[data-baseweb="select"] input,
section[data-testid="stSidebar"] div[data-baseweb="select"] span,
section[data-testid="stSidebar"] div[data-baseweb="tag"],
section[data-testid="stSidebar"] div[data-baseweb="tag"] span {{
  color:#FFFFFF !important;
}}
/* chip nama indikator yang terpilih */
section[data-testid="stSidebar"] div[data-baseweb="tag"] {{
  background-color:rgba(255,255,255,.14);
  border-color:rgba(255,255,255,.32);
}}
section[data-testid="stSidebar"] div[data-baseweb="select"] svg,
section[data-testid="stSidebar"] div[data-baseweb="tag"] svg {{
  fill:#FFFFFF;
}}
</style>
""",
    unsafe_allow_html=True,
)


def kpi_card(label: str, value_html: str, sub_html: str, bar_color: str) -> str:
    return (
        f'<div class="kpi" style="--bar:{bar_color}">'
        f'<div class="lbl">{label}</div>'
        f'<div class="val">{value_html}</div>'
        f'<div class="sub">{sub_html}</div></div>'
    )


def mini_card(name: str, val_html: str, sub_html: str,
              filled: bool, bar_color: str) -> str:
    cls = "mini" if filled else "mini empty"
    return (f'<div class="{cls}" style="--bar:{bar_color}">'
            f'<div class="lbl">{name}</div>'
            f'<div class="val">{val_html}</div>'
            f'<div class="sub">{sub_html}</div></div>')


def section(title: str, note: str = "") -> None:
    note_html = f"<small>{note}</small>" if note else ""
    st.markdown(
        f'<div class="section-title"><span class="dot"></span>'
        f'<h3>{title}</h3>{note_html}</div>',
        unsafe_allow_html=True,
    )


# ===========================================================================
# Grafik
# ===========================================================================
def _base_layout(fig: go.Figure, height: int = 330) -> go.Figure:
    fig.update_layout(
        height=height,
        margin={"l": 8, "r": 18, "t": 40, "b": 8},
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "'Plus Jakarta Sans', sans-serif", "size": 12.5,
              "color": C["ink"]},
        hovermode="x unified",
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02,
                "xanchor": "left", "x": 0},
        title={"font": {"size": 13.5, "color": C["ink"]}, "x": 0.01,
               "xanchor": "left"},
    )
    fig.update_xaxes(showgrid=False, linecolor=C["line"])
    fig.update_yaxes(gridcolor="#ECF1F7", zeroline=False)
    return fig


def _color_traces(fig: go.Figure) -> None:
    for i, tr in enumerate(fig.data):
        col = SERIES_PALETTE[i % len(SERIES_PALETTE)]
        tr.line.color = col
        tr.marker.color = col


def chart_trend(df: pd.DataFrame, *, title: str, kind: str,
                multi_region: bool, as_bar: bool = False) -> go.Figure:
    order = sorted(df.tahun.unique())
    if multi_region:
        fig = px.line(df, x="tahun", y="nilai", color="wilayah",
                      markers=True, category_orders={"tahun": order})
        _color_traces(fig)
    elif as_bar:
        fig = px.bar(df, x="tahun", y="nilai", category_orders={"tahun": order})
        fig.update_traces(marker={"color": C["navy"], "line": {"width": 0},
                                  "cornerradius": 6},
                          customdata=df["nilai"].map(fmt_rp))
        fig.data[0].hovertemplate = "<b>Tahun %{x}</b><br>%{customdata}<extra></extra>"
        mean_val = df["nilai"].dropna().mean()
        if pd.notna(mean_val):
            fig.add_hline(y=mean_val, line_dash="dot", line_color=C["amber"],
                          line_width=2,
                          annotation_text=f"rata-rata {fmt_rp(mean_val)}",
                          annotation_position="top left",
                          annotation_font_color=C["muted"],
                          annotation_font_size=11)
        fig.update_yaxes(separatethousands=True)
    else:
        fig = px.line(df, x="tahun", y="nilai", markers=True,
                      category_orders={"tahun": order})
        if kind == "gk":
            fig.update_traces(line=dict(color=C["navy"], width=3),
                              marker=dict(size=8, color=C["navy"]),
                              customdata=df["nilai"].map(fmt_rp))
            fig.data[0].hovertemplate = "<b>Tahun %{x}</b><br>%{customdata}<extra></extra>"
        else:
            base_col = C["red"] if kind == "ppm" else C["blue"]
            fill_col = ("rgba(226,58,52,.09)" if kind == "ppm"
                        else "rgba(25,118,197,.09)")
            fig.update_traces(line=dict(color=base_col, width=3),
                              marker=dict(size=8, color=base_col),
                              fill="tozeroy", fillcolor=fill_col)
            fig.data[0].hovertemplate = "<b>Tahun %{x}</b><br>%{y:,.2f}<extra></extra>"

    fig.update_layout(title=title)
    if kind == "ppm":
        fig.update_yaxes(ticksuffix="%")
    return _base_layout(fig)


def chart_compare(df: pd.DataFrame, indikator: str, tahun: int) -> go.Figure:
    d = df[(df.indikator == indikator) & (df.tahun == tahun)].copy()
    d = d.dropna(subset=["nilai"]).sort_values("nilai").tail(12)
    is_ppm = kind_of(indikator) == "ppm"
    d["label"] = d.nilai.map(lambda v: fmt_num(v, smart_dec(v)))
    color = C["red"] if is_ppm else C["navy"]
    fig = px.bar(d, x="nilai", y="wilayah", orientation="h", text="label")
    fig.update_traces(marker={"color": color, "cornerradius": 6,
                              "line": {"width": 0}},
                      textposition="outside", cliponaxis=False)
    fig.update_layout(
        title=f"{short_label(indikator)} · {tahun}",
        xaxis_ticksuffix="%" if is_ppm else None,
    )
    if not is_ppm:
        fig.update_xaxes(separatethousands=True)
    fig.update_xaxes(title=None)
    fig.update_yaxes(title=None)
    return _base_layout(fig, height=max(300, 34 * len(d) + 80))


def chart_pie_miskin(ppm_value: float, tahun: int, wilayah_label: str) -> go.Figure:
    """Pie komposisi Penduduk Miskin vs Sejahtera dari Persentase Penduduk
    Miskin (P0). Legenda sengaja hanya menampilkan 'Penduduk Miskin':
    legend bawaan pie dimatikan, lalu ditambah satu entri legend tiruan
    (marker tak kasatmata) dengan nama tersebut."""
    v = max(0.0, min(100.0, float(ppm_value)))
    fig = go.Figure()
    fig.add_trace(go.Pie(
        labels=["Penduduk Sejahtera", "Penduduk Miskin"],
        values=[100.0 - v, v],
        sort=False, direction="clockwise", hole=0,
        marker={"colors": [C["navy"], C["red"]],
                "line": {"color": "#FFFFFF", "width": 2}},
        texttemplate="%{customdata}",
        textposition="inside",
        textfont={"color": "#FFFFFF", "size": 15},
        customdata=[[fmt_num(100 - v, 2) + "%"], [fmt_num(v, 2) + "%"]],
        hovertemplate="<b>%{label}</b><br>%{customdata[0]} dari populasi"
                      "<extra></extra>",
        showlegend=False,
    ))
    # entri legenda tunggal (titik kosong tidak tergambar pada kanvas)
    fig.add_trace(go.Scatter(
        x=[None], y=[None], mode="markers",
        marker={"size": 12, "symbol": "square", "color": C["red"]},
        name="Penduduk Miskin",
        showlegend=True,
    ))
    fig.update_layout(title=f"Komposisi Penduduk · {wilayah_label} · {tahun}")
    fig = _base_layout(fig)
    # Legenda dipindah ke BAWAH pie: posisi bawaan (kiri-atas, satu band
    # dengan judul) menutupi teks "Komposisi Penduduk ...".
    fig.update_layout(
        legend={"orientation": "h", "yanchor": "top", "y": -0.04,
                "xanchor": "left", "x": 0},
        margin={"b": 54},
    )
    return fig


def chart_scatter_ekonomi_tpt(df: pd.DataFrame) -> go.Figure | None:
    """Sebar Pertumbuhan Ekonomi vs TPT per tahun (uji intuitif hukum Okun).
    Garis putus-putus = rata-rata tiap sumbu (pembatas kuadran)."""
    pe = (df[df.indikator == "Pertumbuhan Ekonomi"]
          .dropna(subset=["nilai"]).set_index("tahun")["nilai"])
    tp = (df[df.indikator == "Tingkat Pengangguran Terbuka"]
          .dropna(subset=["nilai"]).set_index("tahun")["nilai"])
    common = pe.index.intersection(tp.index).sort_values()
    if len(common) < 3:
        return None
    x = pe.loc[common].astype(float)
    y = tp.loc[common].astype(float)
    fig = go.Figure(go.Scatter(
        x=x, y=y, mode="markers+text",
        customdata=[int(t) for t in common],
        text=[str(int(t)) for t in common],
        textposition="top center",
        textfont={"size": 11, "color": C["muted"]},
        marker={"size": 14, "color": C["blue"], "opacity": 0.85,
                "line": {"width": 1.5, "color": "#FFFFFF"}},
        hovertemplate=("<b>Tahun %{customdata}</b><br>"
                       "Pertumbuhan ekonomi: %{x:.2f}%<br>"
                       "TPT: %{y:.2f}%<extra></extra>"),
        showlegend=False,
    ))
    fig.add_vline(x=float(x.mean()), line_width=1, line_dash="dot",
                  line_color=C["amber"])
    fig.add_hline(y=float(y.mean()), line_width=1, line_dash="dot",
                  line_color=C["amber"])
    fig.update_xaxes(title="Pertumbuhan Ekonomi (%)")
    fig.update_yaxes(title="Tingkat Pengangguran Terbuka (%)")
    fig = _base_layout(fig)
    fig.update_layout(hovermode="closest")
    return fig


def chart_dual_kemiskinan(df: pd.DataFrame) -> go.Figure | None:
    """Dual-axis: batang % penduduk miskin vs garis garis kemiskinan (Rp).
    Menjawab narasi klasik: P0 turun sementara garis kemiskinan naik."""
    ppm = (df[df.indikator.map(kind_of) == "ppm"]
           .dropna(subset=["nilai"]).set_index("tahun")["nilai"])
    gk = (df[df.indikator.map(kind_of) == "gk"]
          .dropna(subset=["nilai"]).set_index("tahun")["nilai"])
    common = ppm.index.intersection(gk.index).sort_values()
    if len(common) < 2:
        return None
    thn = [int(t) for t in common]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=thn, y=[float(ppm[t]) for t in common],
        name="% Penduduk Miskin",
        marker={"color": "rgba(226,58,52,.72)"},
        hovertemplate="<b>%{x}</b><br>Penduduk miskin: %{y:.2f}%"
                      "<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=thn, y=[float(gk[t]) for t in common],
        name="Garis Kemiskinan (Rp)", yaxis="y2",
        mode="lines+markers",
        line={"color": C["navy"], "width": 3},
        marker={"size": 7, "color": C["navy"],
                "line": {"width": 1.5, "color": "#FFFFFF"}},
        hovertemplate="<b>%{x}</b><br>Garis kemiskinan: Rp%{y:,.0f}"
                      "<extra></extra>",
    ))
    fig.update_layout(
        yaxis={"title": "Persentase Penduduk Miskin", "ticksuffix": "%"},
        yaxis2={"title": "Garis Kemiskinan (Rp)", "overlaying": "y",
                "side": "right", "showgrid": False,
                "tickformat": ",.0f"},
    )
    fig = _base_layout(fig)
    fig.update_layout(hovermode="x unified")
    return fig


def chart_overlay(df: pd.DataFrame, ind_a: str, ind_b: str,
                  chart_type: str = "Garis") -> go.Figure | None:
    """Dua indikator ditumpang-tindih dalam satu kanvas waktu.

    - Kedua deret digambar dalam warna kontras (navy & merah).
    - Bila keduanya satu 'keluarga satuan' (sama-sama persen atau sama-sama
      bukan), dipakai satu sumbu-Y; bila satuannya berbeda, indikator kedua
      otomatis pindah ke sumbu-Y kanan (ala Dinamika Kemiskinan) agar
      skala tetap terbaca.
    - chart_type: 'Garis' | 'Area' | 'Batang' — dapat dipilih user.
    """
    def _series(ind: str) -> pd.DataFrame:
        f = df[df.indikator == ind].dropna(subset=["nilai"])
        return (f.groupby("tahun", as_index=False)["nilai"].mean()
                 .sort_values("tahun"))

    fa, fb = _series(ind_a), _series(ind_b)
    common = sorted(set(fa.tahun) & set(fb.tahun))
    if len(common) < 2:
        return None
    ya = [float(v) for v in fa.set_index("tahun").loc[common, "nilai"]]
    yb = [float(v) for v in fb.set_index("tahun").loc[common, "nilai"]]

    pct_a, pct_b = is_percent_family(ind_a), is_percent_family(ind_b)
    dual = pct_a != pct_b
    lbl_a, lbl_b = short_label(ind_a), short_label(ind_b)
    is_bar = chart_type == "Batang"
    is_area = chart_type == "Area"

    specs = [
        (ya, C["navy"], lbl_a, False, "rgba(10,61,110,.14)", pct_a),
        (yb, C["red"], lbl_b, dual, "rgba(226,58,52,.14)", pct_b),
    ]

    fig = go.Figure()
    for y, col, lbl, on_y2, fill_col, pct in specs:
        yaxis_ref = "y2" if on_y2 else "y"
        hov = f"<b>{lbl}</b>: " + "%{y:,.2f}<extra></extra>"
        if is_bar:
            fig.add_trace(go.Bar(
                x=common, y=y, name=lbl, yaxis=yaxis_ref,
                marker={"color": col, "cornerradius": 5,
                        "opacity": 0.85 if dual else 1.0,
                        "line": {"width": 0}},
                hovertemplate=hov))
        else:
            area_kw = ({"fill": "tozeroy", "fillcolor": fill_col}
                       if is_area else {})
            fig.add_trace(go.Scatter(
                x=common, y=y, name=lbl, yaxis=yaxis_ref,
                mode="lines+markers",
                line={"color": col, "width": 3},
                marker={"size": 7, "color": col,
                        "line": {"width": 1.5, "color": "#FFFFFF"}},
                hovertemplate=hov, **area_kw))

    fig.update_layout(title=f"{lbl_a} vs {lbl_b}")
    if is_bar:
        fig.update_layout(barmode="group")
    if dual:
        fig.update_layout(
            yaxis={"title": lbl_a, "ticksuffix": "%" if pct_a else None,
                   "separatethousands": not pct_a},
            yaxis2={"title": lbl_b, "overlaying": "y", "side": "right",
                    "showgrid": False, "ticksuffix": "%" if pct_b else None,
                    "separatethousands": not pct_b},
        )
    else:
        fig.update_layout(
            yaxis={"ticksuffix": "%" if pct_a else None,
                   "separatethousands": not pct_a},
        )
    # kanvas sedikit lebih tinggi dari default (330 px) agar dua deret
    # yang ditumpangkan tetap terbaca lega
    return _base_layout(fig, height=380)


# --- Gauge IPM dinonaktifkan (dikomentari) ---
# def _kategori_ipm(v: float) -> tuple[str, str]:
#     """(nama kategori, warna batang) menurut ambang resmi BPS."""
#     if v >= 80:
#         return "Sangat Tinggi", C["navy"]
#     if v >= 70:
#         return "Tinggi", "#1F9D66"
#     if v >= 60:
#         return "Sedang", C["amber"]
#     return "Rendah", C["red"]
#
#
# def gauge_ipm(value: float, prev: float | None, label: str,
#               tahun: int) -> go.Figure:
#     """Gauge IPM dengan pita kategori BPS: <60 Rendah, 60–70 Sedang,
#     70–80 Tinggi, >=80 Sangat Tinggi."""
#     ket, warna = _kategori_ipm(value)
#     mode = "gauge+number+delta" if prev is not None else "gauge+number"
#     delta_cfg = None
#     if prev is not None:
#         delta_cfg = {"reference": float(prev), "valueformat": "+.2f"}
#     fig = go.Figure(go.Indicator(
#         mode=mode,
#         value=float(value),
#         number={"valueformat": ".2f", "font": {"size": 32}},
#         delta=delta_cfg,
#         title={"text": f"{label} · {tahun}",
#                "font": {"size": 13, "color": C["ink"]}},
#         gauge={
#             "axis": {"range": [40, 90], "tickvals": [40, 60, 70, 80, 90],
#                      "tickfont": {"size": 10, "color": C["muted"]}},
#             "bar": {"color": warna, "thickness": 0.28},
#             "bgcolor": "rgba(0,0,0,0)",
#             "steps": [
#                 {"range": [40, 60], "color": "rgba(226,58,52,.13)"},
#                 {"range": [60, 70], "color": "rgba(242,169,59,.15)"},
#                 {"range": [70, 80], "color": "rgba(31,157,102,.15)"},
#                 {"range": [80, 90], "color": "rgba(10,61,110,.12)"},
#             ],
#             "threshold": {"line": {"color": C["ink"], "width": 2},
#                           "thickness": 0.75, "value": float(value)},
#         },
#     ))
#     fig.add_annotation(text=f"Kategori: <b>{ket}</b>",
#                        x=0.5, y=-0.06, xref="paper", yref="paper",
#                        showarrow=False,
#                        font={"size": 12.5, "color": C["ink"]})
#     fig.update_layout(
#         height=245,
#         margin={"l": 30, "r": 30, "t": 46, "b": 40},
#         paper_bgcolor="rgba(0,0,0,0)",
#         font={"family": "'Plus Jakarta Sans', sans-serif", "size": 12.5,
#               "color": C["ink"]},
#     )
#     return fig


# ===========================================================================
# Pemuatan otomatis (tiap sesi boot): spreadsheet live -> fallback snapshot
# ===========================================================================
@st.cache_data(ttl=600, show_spinner=False)
def _live_bytes() -> bytes:
    return download_xlsx_bytes(SHEET_XLSX_URL)


# Penggabungan multi-file lokal dihentikan (data hanya dari spreadsheet).
# def _merge_priority(frames: list[pd.DataFrame]) -> pd.DataFrame:
#     """Gabungkan frame dgn urutan prioritas menurun; kunci (w,i,t) sama ->
#     baris dari frame lebih awal yang bernilai menang."""
#     df = pd.concat([f for f in frames if f is not None], ignore_index=True)
#     if df.empty:
#         return df
#     df["_has_val"] = df["nilai"].notna()
#     df = (df.sort_values(["wilayah", "indikator", "tahun", "_has_val"],
#                          ascending=[True, True, True, False], kind="stable")
#             .drop_duplicates(subset=["wilayah", "indikator", "tahun"],
#                              keep="first")
#             .drop(columns="_has_val"))
#     return df.reset_index(drop=True)


@st.cache_data(ttl=300, show_spinner="Memuat data…")
def load_auto() -> tuple[pd.DataFrame, list[str]]:
    """Data diambil HANYA dari Google Spreadsheet live — tanpa penggabungan
    file Excel lokal lain. Gagal jaringan -> pakai salinan snapshot terakhir
    dari spreadsheet yang sama (data/_live_snapshot.xlsx)."""
    try:
        blob = _live_bytes()
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        SYNC_TARGET.write_bytes(blob)
        return load_workbook_df(io.BytesIO(blob)), []
    except Exception:                                   # noqa: BLE001
        # Fallback offline: salinan lokal terakhir dari spreadsheet yang sama.
        if SYNC_TARGET.exists():
            return load_workbook_df(SYNC_TARGET), []
        raise


# Sinkronisasi data: unduhan segar terjadi pada setiap MUAT-ULANG HALAMAN
# (F5 atau klik logo bara-mendoan — keduanya membuka sesi browser baru).
# Rerun widget biasa tetap memakai cache agar tampil responsif.
_first_page_load = "page_loaded" not in st.session_state
st.session_state["page_loaded"] = True
if _first_page_load:
    _live_bytes.clear()
    load_auto.clear()

df_data: pd.DataFrame | None
try:
    df_data, _errs = load_auto()
except Exception:                                       # noqa: BLE001
    df_data = None

if df_data is None or df_data.empty:
    st.error("⏳ Belum ada data yang dapat dimuat. Pastikan koneksi internet "
             "tersedia saat pertama kali membuka dashboard, lalu muat ulang "
             "halaman.")
    st.stop()

# ---------------------------------------------------------------------------
# Normalisasi nama indikator: penamaan lama -> istilah resmi terbaru.
# ---------------------------------------------------------------------------
_RENAME_INDIKATOR = {"Pendapatan Perkapita": "PDRB Perkapita"}


def _norm_indikator(i: str) -> str:
    if i in _RENAME_INDIKATOR:
        return _RENAME_INDIKATOR[i]
    for old, new in _RENAME_INDIKATOR.items():
        if i.lower().startswith(old.lower()):
            # varian dalam kurung (mis. " (Harga Konstan)") tetap dipertahankan
            return new + i[len(old):]
    return i


df_data["indikator"] = df_data["indikator"].map(_norm_indikator)

wilayah_all = sorted(df_data.wilayah.unique())
indikator_all = sorted(df_data.indikator.unique())
tahun_all = sorted(df_data.tahun.unique())

# ===========================================================================
# SIDEBAR
# ===========================================================================
with st.sidebar:
    st.markdown(
        f"""
        <div class="sidebar-brand">
          <div class="logo">{bps_logo(44)}</div>
          <div>
            <b>NewBaramendoan</b>
            <span>BPS Kab. Banjarnegara</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.divider()

    # Wilayah dikunci ke Kabupaten Banjarnegara (opsi pemilihan dihapus).
    kab = next((w for w in wilayah_all if w.lower() == "banjarnegara"),
               wilayah_all[0] if wilayah_all else "")
    wilayah_pick = [kab] if kab else []

    rentang = ((tahun_all[0], tahun_all[-1])
               if len(tahun_all) > 1 else (tahun_all[0], tahun_all[0]))
    if len(tahun_all) > 1:
        rentang = st.select_slider("Rentang Tahun", options=tahun_all,
                                   value=rentang)

    prio = sorted(indikator_all,
                  key=lambda i: {"gk": 0, "ppm": 1}.get(kind_of(i), 2))
    # Default fokus: Indeks Pembangunan Manusia (Laki-laki & Perempuan)
    default_ind = [i for i in indikator_all
                   if i.startswith("Indeks Pembangunan Manusia (")][:2]
    if not default_ind:                     # fallback bila varian IPM absen
        default_ind = [i for i in prio if kind_of(i) in {"gk", "ppm"}][:2] \
            or prio[:2]
    indikator_pick = st.multiselect(
        "Indikator Ditampilkan", prio, default=default_ind,
        help="Kartu fokus dan grafik utama mengikuti pilihan ini.",
    )
    indikator_pick = indikator_pick or default_ind[:1]

multi_region = False        # wilayah terkunci satu kabupaten

# ===========================================================================
# BANNER JUDUL (logo + fokus dinamis)
# ===========================================================================
judul_wilayah = ", ".join(wilayah_pick[:3]) + (
    f" +{len(wilayah_pick) - 3}" if len(wilayah_pick) > 3 else "")
fokus_labels = [short_label(i) for i in indikator_pick]
fokus_text = ", ".join(fokus_labels[:4]) + ("…" if len(fokus_labels) > 4 else "")
st.markdown(
    f"""
    <div class="hero">
      <div class="hero-logos">
        <a class="hero-logo-plain" href="?" title="Muat ulang halaman (ambil data terbaru)"
           aria-label="Muat ulang halaman">{mendoan_logo(74)}</a>
        <a class="hero-logo-plain" href="https://banjarnegarakab.bps.go.id/id"
           target="_blank" rel="noopener" title="BPS Kabupaten Banjarnegara"
           aria-label="BPS Kabupaten Banjarnegara">{bps_logo(74)}</a>
      </div>
      <div class="hero-main">
        <div class="hero-text">
          <h1>NewBaramendoan</h1>
          <p class="subjudul">New Banjarnegara Mencari Data tanpa Dolan</p>
          <p><b>{fokus_text}</b></p>
          <p class="fokus">Periode {rentang[0]}–{rentang[-1]}</p>
        </div>
        <div class="hero-side">
          <div class="hero-stamp">{data_updated_at():%d %b %Y %H:%M}</div>
          <div class="badge-stack">
            <span>Indikator difokuskan : {len(indikator_pick)}</span>
            <span>Data Statistik : {len(CANONICAL)}</span>
          </div>
        </div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Kartu KPI fokus (dinamis mengikuti indikator terpilih)
# ---------------------------------------------------------------------------
def latest_stat(frame: pd.DataFrame):
    if frame.empty:
        return None
    y_last = frame.tahun.max()
    vs = frame.loc[frame.tahun == y_last, "nilai"].dropna()
    if vs.empty:
        return None
    v_last = float(vs.iloc[-1])
    prev_years = frame[frame.tahun < y_last].tahun
    if prev_years.empty:
        return y_last, v_last, None, None
    y_prev = prev_years.max()
    vps = frame.loc[frame.tahun == y_prev, "nilai"].dropna()
    v_prev = float(vps.iloc[-1]) if not vps.empty else None
    return y_last, v_last, y_prev, v_prev


cards = []
sub_frames = {
    ind: df_data[df_data.indikator == ind]
    .sort_values("tahun") for ind in indikator_pick[:2]
}
for idx, ind in enumerate(indikator_pick[:2]):
    frame = sub_frames[ind]
    stat = latest_stat(frame)
    if stat is None:
        continue
    y, v, yp, vp = stat
    d, unit = delta_display(v, vp, ind)
    dec = smart_dec(v)
    suffix = "%" if is_percent_family(ind) else ""
    cards.append(kpi_card(
        short_label(ind),
        f"{fmt_num(v, dec)}{suffix}",
        delta_chip(d, unit=unit) +
        f'<span class="hint">tahun {y}' + (f" vs {yp}" if yp else "") + "</span>",
        FOCUS_BAR_COLORS[idx % len(FOCUS_BAR_COLORS)],
    ))

df_f = df_data[
    df_data.wilayah.isin(wilayah_pick) & df_data.tahun.between(*rentang)
].copy()

n_obs = len(df_f.dropna(subset=["nilai"]))
cards.append(kpi_card(
    "Observasi Terfilter",
    f"{fmt_num(n_obs, 0)}",
    f'<span class="chip chip-info">{judul_wilayah}</span>'
    f'<span class="hint">{rentang[0]}–{rentang[-1]}</span>',
    C["green"],
))
# Periode mengikuti indikator terpilih (bukan seluruh workbook)
thn_pick = sorted(df_data.loc[df_data.indikator.isin(indikator_pick),
                              "tahun"].dropna().unique().tolist())
if thn_pick:
    cards.append(kpi_card(
        "Periode Tersedia",
        f"{thn_pick[0]}–{thn_pick[-1]}",
        f'<span class="chip chip-info">{len(thn_pick)} titik tahun</span>'
        f'<span class="hint">{len(indikator_pick)} indikator terpilih</span>',
        C["blue"],
    ))
else:
    cards.append(kpi_card(
        "Periode Tersedia", "-",
        '<span class="hint">indikator terpilih belum punya data</span>',
        C["blue"],
    ))

st.markdown('<div class="kpi-grid">' + "".join(cards) + "</div>",
            unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Komposisi penduduk: miskin vs sejahtera (dari Persentase Penduduk Miskin).
# Hanya tampil saat 'Garis Kemiskinan' termasuk indikator yang dipilih user.
# ---------------------------------------------------------------------------
show_pie = any(kind_of(i) == "gk" for i in indikator_pick)
if show_pie:
    ppm_frame = (df_f[df_f.indikator.map(kind_of) == "ppm"]
                 .dropna(subset=["nilai"]))
    section("Komposisi Penduduk", "penduduk miskin vs sejahtera")
    if ppm_frame.empty:
        st.info("Persentase Penduduk Miskin tidak tersedia pada filter ini.")
    else:
        if multi_region:                    # rata-rata antar wilayah terpilih
            ppm_per_tahun = ppm_frame.groupby("tahun")["nilai"].mean()
        else:
            ppm_per_tahun = ppm_frame.set_index("tahun")["nilai"].sort_index()
        y_pie = int(ppm_per_tahun.index.max())
        ppm_val = float(ppm_per_tahun.loc[y_pie])
        with st.container(border=True):
            st.plotly_chart(
                chart_pie_miskin(ppm_val, y_pie, judul_wilayah),
                width="stretch", config={"displayModeBar": False})

# ---------------------------------------------------------------------------
# Dual-axis dinamika kemiskinan: % penduduk miskin vs garis kemiskinan.
# Tampil saat varian Tingkat Kemiskinan (GK / %P0) termasuk pilihan.
# ---------------------------------------------------------------------------
show_kemiskinan = any(kind_of(i) in {"gk", "ppm"} for i in indikator_pick)
if show_kemiskinan:
    dual_fig = chart_dual_kemiskinan(df_f)
    if dual_fig is not None:
        section("Dinamika Kemiskinan",
                "% penduduk miskin (batang) vs garis kemiskinan Rp (garis)")
        with st.container(border=True):
            st.plotly_chart(dual_fig, width="stretch",
                            config={"displayModeBar": False})

# ---------------------------------------------------------------------------
# Gauge IPM vs kategori resmi BPS — DINONAKTIFKAN (dikomentari)
# ---------------------------------------------------------------------------
# show_gauge = any(i.startswith("Indeks Pembangunan Manusia")
#                  for i in indikator_pick)
# if show_gauge:
#     section("Kategori IPM", "ambang BPS: 60 Sedang · 70 Tinggi · 80 Sangat Tinggi")
#     ipm_vars: list[tuple[str, pd.DataFrame]] = []
#     for var_label in ("Indeks Pembangunan Manusia (Laki-laki)",
#                       "Indeks Pembangunan Manusia (Perempuan)"):
#         v = (df_f[df_f.indikator == var_label]
#              .dropna(subset=["nilai"]).sort_values("tahun"))
#         if not v.empty:
#             ipm_vars.append((var_label, v))
#     if ipm_vars:
#         cols = st.columns(len(ipm_vars))
#         for col, (nm, v) in zip(cols, ipm_vars):
#             lbl = ("IPM Laki-laki" if "laki" in nm.lower()
#                    else "IPM Perempuan" if "perempuan" in nm.lower()
#                    else nm[:24])
#             prev = float(v.nilai.iloc[-2]) if len(v) > 1 else None
#             with col:
#                 with st.container(border=True):
#                     st.plotly_chart(
#                         gauge_ipm(float(v.nilai.iloc[-1]), prev, lbl,
#                                   int(v.tahun.iloc[-1])),
#                         width="stretch",
#                         config={"displayModeBar": False})

# ---------------------------------------------------------------------------
# Scatter pertumbuhan ekonomi vs pengangguran
# ---------------------------------------------------------------------------
sc_fig = chart_scatter_ekonomi_tpt(df_f)
if sc_fig is not None:
    section("Ekonomi vs Pengangguran",
            "sebaran tahunan · garis putus-putus = rata-rata tiap sumbu")
    with st.container(border=True):
        st.plotly_chart(sc_fig, width="stretch",
                        config={"displayModeBar": False})

# ---------------------------------------------------------------------------
# Grid 21 indikator kanonik (termasuk yang belum terisi)
# ---------------------------------------------------------------------------
df_reg = df_data[df_data.wilayah.isin(wilayah_pick)]
canon_filled = 0
mini_cards_html = []


def _short_var(nm: str) -> str:
    """Label pendek untuk varian indikator pada kartu."""
    low = (nm or "").lower()
    if "garis" in low:
        return "Garis kemiskinan"
    if "persentase" in low or "(persen)" in low:
        return "% penduduk miskin"
    if "laki" in low:
        return "Laki-laki"
    if "perempuan" in low:
        return "Perempuan"
    m = re.search(r"\(([^)]+)\)", nm)
    return m.group(1) if m else nm[:24]


for pos, name in enumerate(CANONICAL):
    sub = df_reg[df_reg.indikator.map(canon_key) == name]
    valid = sub.dropna(subset=["nilai"])
    bar = SERIES_PALETTE[pos % len(SERIES_PALETTE)]
    if valid.empty:
        mini_cards_html.append(mini_card(
            name, "belum tersedia", "", False, bar))
        continue
    canon_filled += 1
    y_last = valid.tahun.max()
    rows = valid[valid.tahun == y_last].sort_values("indikator")
    n_var = rows.indikator.nunique()

    # --- tepat 2 varian -> dua mode kartu ---
    if n_var == 2:
        if name == "Tingkat Kemiskinan":
            # perilaku lama: kartu dua-wajah dengan toggle ⇄ (GK <-> %P0)
            faces = []
            for _, r in rows.iterrows():
                suf = "%" if "(persen)" in r.indikator.lower() else ""
                val = f"{fmt_num(r.nilai, smart_dec(r.nilai))}{suf}"
                subh = (f'<span class="chip chip-year">{y_last}</span>'
                        f"<span>{_short_var(r.indikator)}</span>")
                faces.append((val, subh))
            cb = f'<input type="checkbox" id="vt{pos}" class="vswitch">'
            face_html = "".join(
                f'<div class="face f{k}"><div class="val">{v}</div>'
                f'<div class="sub">{s}</div></div>'
                for k, (v, s) in enumerate(faces))
            mini_cards_html.append(
                f'<div class="mini" style="--bar:{bar}">{cb}'
                f'<div class="lbl">{name}</div>{face_html}'
                f'<label class="vtoggle" for="vt{pos}">⇄ 2 varian</label></div>')
            continue
        # varian lain (mis. IPM L/P): kedua nilai tampil berdampingan
        halves = []
        for _, r in rows.iterrows():
            suf = "%" if "(persen)" in r.indikator.lower() else ""
            val = f"{fmt_num(r.nilai, smart_dec(r.nilai))}{suf}"
            halves.append(
                f'<div class="duo-item"><div class="val">{val}</div>'
                f'<div class="duo-lbl">{_short_var(r.indikator)}</div></div>')
        sub_html = f'<span class="chip chip-year">{y_last}</span>'
        mini_cards_html.append(
            f'<div class="mini" style="--bar:{bar}">'
            f'<div class="lbl">{name}</div>'
            f'<div class="duo">{"".join(halves)}</div>'
            f'<div class="sub">{sub_html}</div></div>')
        continue

    # --- satu / lebih dari dua varian -> tampilan gabungan ---
    parts: list[str] = []
    extra_sub = ""
    for _, r in rows.iterrows():
        suf = "%" if "(persen)" in r.indikator.lower() else ""
        parts.append(f"{fmt_num(r.nilai, smart_dec(r.nilai))}{suf}")
    if name == "Tingkat Kemiskinan":
        gk_rows = rows[rows.indikator.map(kind_of) == "gk"]
        if not gk_rows.empty:
            extra_sub = f"Garis kemiskinan {fmt_rp(float(gk_rows.nilai.iloc[0]))}"
    val_html = " · ".join(parts)
    sub_html = f'<span class="chip chip-year">{y_last}</span>'
    if extra_sub:
        sub_html += f"<span>{extra_sub}</span>"
    if n_var > 2:
        sub_html += f"<span>({n_var} varian)</span>"
    mini_cards_html.append(mini_card(name, val_html, sub_html, True, bar))

section("Data statistik",
        f"{canon_filled} dari {len(CANONICAL)} sudah memiliki data")
st.markdown('<div class="mini-grid">' + "".join(mini_cards_html) + "</div>",
            unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Grafik utama (dinamis: dua indikator teratas yang dipilih)
# ---------------------------------------------------------------------------
section("Tren Waktu", "mengikuti indikator terpilih")

col_a, col_b = st.columns(2, gap="medium")
for col, ind, empty_msg in (
        (col_a, indikator_pick[0], "Pilih minimal satu indikator di sidebar."),
        (col_b, indikator_pick[1] if len(indikator_pick) > 1 else None,
         "Pilih indikator kedua untuk perbandingan berdampingan."),
):
    with col:
        if ind is None:
            st.info(empty_msg)
            continue
        odf = df_f[df_f.indikator == ind].sort_values("tahun")
        if odf.dropna(subset=["nilai"]).empty:
            st.info(f"Tidak ada data {short_label(ind)} pada filter ini.")
            continue
        k = kind_of(ind)
        with st.container(border=True):
            st.plotly_chart(
                chart_trend(odf, title=short_label(ind), kind=k,
                            multi_region=multi_region,
                            as_bar=(not multi_region and k == "gk")),
                width="stretch", config={"displayModeBar": False})

# ---------------------------------------------------------------------------
# Indikator lainnya (sisanya dari pilihan)
# ---------------------------------------------------------------------------
rest_all = indikator_pick[2:]
if rest_all:
    section("Indikator Lainnya", "dari daftar indikator terpilih")
    pick_other = st.selectbox("Indikator", rest_all,
                              label_visibility="collapsed")
    odf = df_f[df_f.indikator == pick_other].sort_values("tahun")
    if not odf.dropna(subset=["nilai"]).empty:
        with st.container(border=True):
            st.plotly_chart(
                chart_trend(odf, title=short_label(pick_other),
                            kind=kind_of(pick_other),
                            multi_region=multi_region),
                width="stretch", config={"displayModeBar": False})

    # --- grafik tumpang-tindih dua indikator (jenis grafik dapat dipilih) ---
    ov_opts = rest_all if len(rest_all) >= 2 else indikator_pick
    if len(ov_opts) >= 2:
        section("Perbandingan Tumpang-tindih",
                "dua indikator dalam satu kanvas · satuan berbeda otomatis "
                "dipindah ke sumbu kanan")
        sel_a, sel_b, sel_t = st.columns([2, 2, 1])
        with sel_a:
            ov_a = st.selectbox("Indikator 1", ov_opts, key="ov_ind_1")
        with sel_b:
            b_opts = [o for o in ov_opts if o != ov_a] or ov_opts
            ov_b = st.selectbox("Indikator 2", b_opts, key="ov_ind_2")
        with sel_t:
            ov_type = st.selectbox("Jenis Grafik",
                                   ["Garis", "Area", "Batang"],
                                   key="ov_type")
        ov_fig = chart_overlay(df_f, ov_a, ov_b, ov_type)
        if ov_fig is None:
            st.info("Kedua indikator belum punya minimal 2 tahun data yang "
                    "sama untuk dibandingkan.")
        else:
            with st.container(border=True):
                st.plotly_chart(ov_fig, width="stretch",
                                config={"displayModeBar": False})

# ---------------------------------------------------------------------------
# Perbandingan antar wilayah
# ---------------------------------------------------------------------------
if multi_region:
    latest_year = int(df_f.tahun.max())
    cand = next((i for i in indikator_pick
                 if not df_f[(df_f.indikator == i)
                             & (df_f.tahun == latest_year)]
                 .dropna(subset=["nilai"]).empty),
                indikator_pick[0])
    section("Perbandingan Antar Wilayah", f"kondisi terbaru · {latest_year}")
    with st.container(border=True):
        st.plotly_chart(chart_compare(df_f, cand, latest_year),
                        width="stretch", config={"displayModeBar": False})

# ---------------------------------------------------------------------------
# Tabel & unduhan
# ---------------------------------------------------------------------------
section("Data Rinci", "sesuai filter aktif")

# Tabel hanya menampilkan indikator yang dipilih di sidebar,
# konsisten dengan kartu KPI & grafik tren.
tbl_num = df_f[df_f.indikator.isin(indikator_pick)][
    ["wilayah", "indikator", "tahun", "nilai"]].rename(columns={
    "wilayah": "Wilayah", "indikator": "Indikator",
    "tahun": "Tahun", "nilai": "Nilai"})
tbl_show = tbl_num.copy()
tbl_show["Nilai"] = tbl_num.Nilai.map(lambda v: "-" if pd.isna(v)
                                      else fmt_num(v, smart_dec(v)))
tbl_show["Indikator"] = tbl_show.Indikator.map(short_label)
tbl_show = tbl_show.sort_values(["Indikator", "Tahun", "Wilayah"])

st.dataframe(tbl_show, width="stretch", hide_index=True,
             height=min(460, max(220, 38 * len(tbl_show))),
             column_config={
                 "Tahun": st.column_config.NumberColumn(format="%d"),
             })

csv_bytes = tbl_num.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Unduh data terfilter (CSV)", data=csv_bytes,
                   file_name=f"data_statistik_{rentang[0]}-{rentang[-1]}.csv",
                   mime="text/csv")

# ---------------------------------------------------------------------------
# Footer: media sosial BPS (warna senada banner/header)
# ---------------------------------------------------------------------------
_SOC = [
    ("Facebook",
     "https://www.facebook.com/profile.php?id=100068940708645",
     '<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">'
     '<path d="M13.5 21v-7h2.4l.4-3h-2.8V9.1c0-.9.3-1.5 1.6-1.5h1.3V4.9'
     'c-.3 0-1.1-.1-2.1-.1-2.1 0-3.6 1.3-3.6 3.7V11H8.3v3h2.4v7h2.8z"/></svg>'),
    ("Instagram",
     "https://www.instagram.com/bpsbanjarnegara",
     '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" '
     'stroke="currentColor" stroke-width="2">'
     '<rect x="3" y="3" width="18" height="18" rx="5"/>'
     '<circle cx="12" cy="12" r="4"/>'
     '<circle cx="17.2" cy="6.8" r="1" fill="currentColor" stroke="none"/></svg>'),
    ("X (Twitter)",
     "https://twitter.com/bpsbanjarnegara",
     '<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">'
     '<path d="M17.5 3h3.1l-6.8 7.8L21.8 21h-6.3l-4.9-6.4L5 21H1.9l7.3-8.3'
     'L2.2 3h6.4l4.4 5.9L17.5 3zm-1.1 16.2h1.7L7.7 4.7H5.9l10.5 14.5z"/></svg>'),
    ("YouTube",
     "https://www.youtube.com/@bpsbanjarnegara3304",
     '<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">'
     '<path d="M21.6 7.2a2.5 2.5 0 0 0-1.8-1.8C18.2 5 12 5 12 5s-6.2 0-7.8.4'
     'A2.5 2.5 0 0 0 2.4 7.2 26 26 0 0 0 2 12a26 26 0 0 0 .4 4.8 2.5 2.5 0 0 0 '
     '1.8 1.8C5.8 19 12 19 12 19s6.2 0 7.8-.4a2.5 2.5 0 0 0 1.8-1.8A26 26 0 0 0 '
     '22 12a26 26 0 0 0-.4-4.8zM10 15V9l5.2 3L10 15z"/></svg>'),
]
_soc_links = "".join(
    f'<a class="soc" href="{url}" target="_blank" rel="noopener noreferrer">'
    f'{icon}{name}</a>'
    for name, url, icon in _SOC
)
st.markdown(
    f"""
    <div class="site-footer">
      <div class="foot-cols">
        <div class="foot-col">
          <div class="foot-title">Media Sosial</div>
          <div class="soc-grid">{_soc_links}</div>
        </div>
        <div class="foot-col">
          <div class="foot-title">Kontak</div>
          <div class="contact">
            Jl. Selamanik 33 Banjarnegara 53415 Provinsi Jawa Tengah<br/>
            Telp (62-286) 591893<br/>
            Faks (62-286) 592816<br/>
            HP 085195923304<br/>
            Mailbox : <a href="mailto:bps3304@bps.go.id">bps3304@bps.go.id</a>
          </div>
        </div>
      </div>
      <div class="footnote-line">© Badan Pusat Statistik ·
        NewBaramendoan</div>
    </div>
    """,
    unsafe_allow_html=True,
)
