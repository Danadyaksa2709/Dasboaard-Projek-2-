"""
DASHBOARD PROJEK 2
Segmentasi Kecamatan Berdasarkan Ketersediaan Fasilitas Pendidikan dan Kesehatan
Kabupaten Tuban 2022–2025 · Magang MBKM BPS Kabupaten Tuban

CARA MENJALANKAN
  A. VS Code   : buka file ini lalu klik ▶ "Run Python File". Dashboard terbuka otomatis di browser.
  B. Terminal  : python -m streamlit run dashboard_projek2.py
  C. Streamlit Community Cloud: unggah file ini + requirements.txt (+ folder .streamlit/config.toml) ke GitHub.

DATA
  Dashboard mencari file "Data_Projek_2_Python.xlsx" / "Data Projek 2 Python.xlsx" di folder file ini.
  Jika tidak ada, dipakai salinan data yang tertanam di dalam file ini, jadi dashboard tetap jalan.

Library: streamlit, pandas, numpy, plotly, scikit-learn, openpyxl
"""
import base64
import io
import re
import sys
from pathlib import Path

if __name__ == "__main__":
    try:
        from streamlit import runtime as _rt
        _jalan = _rt.exists()
    except Exception:
        _jalan = True
    if not _jalan:
        import subprocess
        print("Menjalankan dashboard... (tekan Ctrl + C untuk berhenti)")
        sys.exit(subprocess.call([sys.executable, "-m", "streamlit", "run", str(Path(__file__).resolve()),
                                  "--theme.base", "light"]))

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler

st.set_page_config(page_title="Dashboard Fasilitas Kecamatan · Kabupaten Tuban", page_icon="🏫", layout="wide")

# =============================================================================================
# 1. PENGATURAN DAN ISTILAH
# =============================================================================================
LINK_PETA = "https://bit.ly/PetaSegmentasiFasilitasKabTuban"
DATA_PATH = r""          # opsional: isi lokasi file Excel jika tidak ditemukan otomatis
SEED, K = 42, 4
BASE = Path(__file__).resolve().parent

FITUR = ["Rasio_Sekolah_SD", "Rasio_Sekolah_SMA", "Rasio_Murid_Guru_SD", "Rasio_Murid_Guru_SMP",
         "Rasio_Murid_Guru_SMA", "Rasio_Puskesmas", "Persen_Desa_Apotek"]

# Setiap indikator ditampilkan dalam ukuran yang mudah dibayangkan orang awam.
#   "per_fasilitas": 1 fasilitas melayani berapa penduduk (10.000 ÷ rasio)  -> makin KECIL makin baik
#   "murid_guru"   : 1 guru mengajar berapa murid                         -> makin KECIL makin baik
#   "persen"       : persen desa yang punya apotek                        -> makin BESAR makin baik
IND = {
    "Rasio_Sekolah_SD": dict(jenis="per_fasilitas", nama="SD/MI", ikon="🏫", kelompok="Pendidikan",
                             ukuran="Penduduk per 1 SD/MI", satuan="penduduk",
                             kalimat="1 SD/MI melayani ± {v} penduduk"),
    "Rasio_Sekolah_SMA": dict(jenis="per_fasilitas", nama="SMA/SMK/MA", ikon="🎓", kelompok="Pendidikan",
                              ukuran="Penduduk per 1 SMA/SMK/MA", satuan="penduduk",
                              kalimat="1 SMA/SMK/MA melayani ± {v} penduduk"),
    "Rasio_Murid_Guru_SD": dict(jenis="murid_guru", nama="Guru SD/MI", ikon="👩‍🏫", kelompok="Pendidikan",
                                ukuran="Murid per 1 guru SD/MI", satuan="murid",
                                kalimat="1 guru SD/MI mengajar ± {v} murid"),
    "Rasio_Murid_Guru_SMP": dict(jenis="murid_guru", nama="Guru SMP/MTs", ikon="👨‍🏫", kelompok="Pendidikan",
                                 ukuran="Murid per 1 guru SMP/MTs", satuan="murid",
                                 kalimat="1 guru SMP/MTs mengajar ± {v} murid"),
    "Rasio_Murid_Guru_SMA": dict(jenis="murid_guru", nama="Guru SMA/SMK/MA", ikon="🧑‍🏫", kelompok="Pendidikan",
                                 ukuran="Murid per 1 guru SMA/SMK/MA", satuan="murid",
                                 kalimat="1 guru SMA/SMK/MA mengajar ± {v} murid"),
    "Rasio_Puskesmas": dict(jenis="per_fasilitas", nama="Puskesmas", ikon="🏥", kelompok="Kesehatan",
                            ukuran="Penduduk per 1 puskesmas", satuan="penduduk",
                            kalimat="1 puskesmas melayani ± {v} penduduk"),
    "Persen_Desa_Apotek": dict(jenis="persen", nama="Apotek", ikon="💊", kelompok="Kesehatan",
                               ukuran="Desa yang punya apotek", satuan="% desa",
                               kalimat="{v}% desa sudah punya apotek"),
}
PILIHAN = {f"{v['ikon']} {v['ukuran']}": f for f, v in IND.items()}

PK, PM, PP, ST = "Pusat Layanan Perkotaan", "Pendidikan Menengah Kuat", "Fasilitas Padat Pengguna", "SMA Terbatas"
URUT = [PK, PM, PP, ST]
# Warna segmen dipilih cukup gelap agar teks putih di atasnya tetap terbaca (kontras ≥ 4,5 : 1)
WARNA = {PK: "#B03A2E", PM: "#A85400", PP: "#0E7C66", ST: "#1F4E79"}
MUDA = {PK: "#FBEAE8", PM: "#FDF0E2", PP: "#E3F4EF", ST: "#E6EEF7"}
CIRI = {
    PK: "Kecamatan kota yang padat. Apotek paling mudah dijumpai, tetapi satu guru harus mengajar murid paling banyak dan satu SD melayani penduduk paling banyak.",
    PM: "Sekolah SMP dan SMA paling banyak dibanding jumlah penduduknya. Kekurangannya, satu puskesmas harus melayani penduduk paling banyak.",
    PP: "Menampung hampir separuh penduduk Tuban. Fasilitasnya mendekati rata-rata, tetapi sekolah SD-nya kurang sehingga satu sekolah dipakai banyak orang.",
    ST: "SMA/SMK/MA paling sedikit dibanding jumlah penduduknya. Sebaliknya, puskesmasnya paling mencukupi dan beban guru SD paling ringan.",
}
SARAN = {
    PK: "Tambah guru SD dan SMP agar beban mengajar lebih ringan.",
    PM: "Tambah puskesmas pembantu atau poskesdes.",
    PP: "Tambah sekolah/ruang kelas SD dan SMA, perluas layanan kesehatan desa.",
    ST: "Prioritaskan pembangunan SMA/SMK/MA atau bantuan transportasi sekolah.",
}
NAVY, TEKS, ABU, GARIS, LATAR = "#0F2A4A", "#1E2B3A", "#4A5664", "#DCE3EC", "#F4F6F9"
HIJAU_T, HIJAU_B = "#1B5E20", "#E3F4E8"
MERAH_T, MERAH_B = "#9B1C1C", "#FDE8E8"
NETRAL_T, NETRAL_B = "#37474F", "#ECEFF3"


# =============================================================================================
# 2. FUNGSI BANTU
# =============================================================================================
def _norm(t):
    return re.sub(r"[^a-z0-9]", "", str(t).lower())


def folder_cari():
    hasil = []
    for awal in [BASE, Path.cwd().resolve()]:
        for f in [awal, awal.parent] + ([s for s in awal.iterdir() if s.is_dir()] if awal.is_dir() else []):
            if f not in hasil:
                hasil.append(f)
    return hasil


def data_valid(d):
    return {"Kecamatan", "Tahun", "Jumlah_Penduduk"}.issubset(d.columns) and set(FITUR).issubset(d.columns)


def baca(sumber):
    d = pd.read_excel(sumber)
    d = d[pd.to_numeric(d["Tahun"], errors="coerce").notna()].copy()
    d["Tahun"] = d["Tahun"].astype(int)
    return d


@st.cache_data(show_spinner=False)
def muat_data():
    kandidat = [Path(DATA_PATH)] if DATA_PATH else []
    for f in folder_cari():
        try:
            kandidat += sorted(p for p in f.glob("*.xls*") if not p.name.startswith("~$")
                               and _norm(p.stem).startswith("dataprojek2"))
        except OSError:
            pass
    for p in kandidat:
        try:
            d = baca(p)
            if data_valid(d):
                return d
        except Exception:
            continue
    return baca(io.BytesIO(base64.b64decode(_DATA_CADANGAN)))


def logo_src():
    for f in folder_cari():
        try:
            for p in sorted(f.iterdir()):
                if p.suffix.lower() in (".png", ".jpg", ".jpeg") and "logo" in _norm(p.stem):
                    return f"data:image/{p.suffix.lower().strip('.')};base64," + base64.b64encode(p.read_bytes()).decode()
        except OSError:
            continue
    return "data:image/png;base64," + _LOGO_CADANGAN


def _segmen_kmeans(rata):
    scaler = StandardScaler().fit(rata[FITUR])
    km = KMeans(n_clusters=K, n_init=200, random_state=SEED).fit(scaler.transform(rata[FITUR]))
    prof = pd.DataFrame(rata[FITUR].values, index=rata.index, columns=FITUR).groupby(km.labels_).mean()
    nama = {}
    for c, b in prof.iterrows():           # nama segmen diberikan dari ciri profilnya (sama seperti di notebook)
        if b["Persen_Desa_Apotek"] == prof["Persen_Desa_Apotek"].max():
            nama[c] = PK
        elif b["Rasio_Sekolah_SMA"] == prof["Rasio_Sekolah_SMA"].max():
            nama[c] = PM
        elif b["Rasio_Sekolah_SMA"] == prof["Rasio_Sekolah_SMA"].min():
            nama[c] = ST
        else:
            nama[c] = PP
    return pd.Series(km.labels_, index=rata.index).map(nama), scaler, km, nama


@st.cache_data(show_spinner=False)
def hitung(d):
    rata = d.groupby("Kecamatan")[FITUR].mean()
    rata["Segmen"], scaler, km, nama = _segmen_kmeans(rata)
    # Segmen jika dihitung dengan data satu tahun saja (dipetakan ke pusat segmen yang sama)
    th = d[["Kecamatan", "Tahun"] + FITUR].copy()
    th["Segmen"] = pd.Series(km.predict(scaler.transform(th[FITUR]))).map(nama).values
    # Uji kestabilan: kelompokkan ulang dengan data tiap tahun (2022, 2023, 2024, 2025), tanpa 2024, tanpa Tuban
    skenario = {f"Hanya {t}": d[d["Tahun"] == t] for t in sorted(d["Tahun"].unique())}
    skenario[f"Tanpa {sorted(d['Tahun'].unique())[-2]}"] = d[d["Tahun"] != sorted(d["Tahun"].unique())[-2]]
    skenario["Tanpa Kecamatan Tuban"] = d[d["Kecamatan"] != "Tuban"]
    tetap = pd.DataFrame(index=rata.index)
    for nm, sub in skenario.items():
        r2 = sub.groupby("Kecamatan")[FITUR].mean()
        seg2 = _segmen_kmeans(r2)[0]
        tetap[nm] = (seg2 == rata.loc[r2.index, "Segmen"]).reindex(rata.index)
    rata["Stabil_%"] = tetap.mean(axis=1, skipna=True) * 100
    return rata, th


def ubah(f, x):
    """Ubah rasio ke ukuran yang mudah dibayangkan."""
    if x is None or pd.isna(x):
        return np.nan
    return 10000 / x if IND[f]["jenis"] == "per_fasilitas" else x


def kecil_baik(f):
    return IND[f]["jenis"] != "persen"


def bulat(f, v):
    if pd.isna(v):
        return np.nan
    if IND[f]["jenis"] == "per_fasilitas":
        return round(v / 100) * 100
    return round(v, 1) if IND[f]["jenis"] == "murid_guru" else round(v, 1)


def fmt(v, d=0):
    if v is None or pd.isna(v) or np.isinf(v):
        return "–"
    return f"{v:,.{d}f}".replace(",", "#").replace(".", ",").replace("#", ".")


def teks_nilai(f, v):
    """Nilai indikator dalam satuan mudah, misal '1.300 penduduk' atau '12,4 murid'."""
    b = bulat(f, v)
    j = IND[f]["jenis"]
    return f"{fmt(b)} penduduk" if j == "per_fasilitas" else (f"{fmt(b, 1)} murid" if j == "murid_guru" else f"{fmt(b, 1)}%")


def kalimat(f, v):
    b = bulat(f, v)
    return IND[f]["kalimat"].format(v=fmt(b) if IND[f]["jenis"] == "per_fasilitas" else fmt(b, 1))


def status(f, v, ref):
    """Bandingkan dengan rata-rata kabupaten: lebih baik / setara / lebih kurang (batas ±10%)."""
    if pd.isna(v) or pd.isna(ref) or ref == 0:
        return "Setara rata-rata"
    rasio = v / ref
    if kecil_baik(f):
        return "Lebih baik" if rasio <= 0.9 else ("Lebih kurang" if rasio >= 1.1 else "Setara rata-rata")
    return "Lebih baik" if rasio >= 1.1 else ("Lebih kurang" if rasio <= 0.9 else "Setara rata-rata")


def chip(s):
    t, b, ik = {"Lebih baik": (HIJAU_T, HIJAU_B, "▲"), "Lebih kurang": (MERAH_T, MERAH_B, "▼"),
                "Setara rata-rata": (NETRAL_T, NETRAL_B, "●")}[s]
    return f"<span class='chip' style='color:{t};background:{b}'>{ik} {s}</span>"


def seg_label(s, kecil=False):
    return f"<span class='seg{' kecil' if kecil else ''}' style='background:{WARNA[s]}'>{s}</span>"


try:
    _VER = tuple(int(v) for v in st.__version__.split(".")[:2])
except Exception:
    _VER = (1, 0)


def grafik(fig):
    cfg = {"displaylogo": False, "modeBarButtonsToRemove": ["select2d", "lasso2d", "autoScale2d"],
           "toImageButtonOptions": {"format": "png", "scale": 2}}
    if _VER >= (1, 50):
        st.plotly_chart(fig, width="stretch", config=cfg)
    else:
        st.plotly_chart(fig, use_container_width=True, config=cfg)


def gaya(fig, tinggi=380):
    fig.update_layout(template="plotly_white", height=tinggi, separators=",.",
                      font=dict(family="Inter, Segoe UI, Arial, sans-serif", size=14, color=TEKS),
                      margin=dict(l=10, r=30, t=10, b=10), paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
                      hoverlabel=dict(bgcolor="#FFFFFF", font_size=14, font_color=TEKS, bordercolor=GARIS),
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                                  font=dict(size=13, color=TEKS), bgcolor="rgba(0,0,0,0)"))
    fig.update_xaxes(gridcolor="#EDF1F5", zeroline=False, linecolor=GARIS, tickfont=dict(color=TEKS, size=13),
                     title_font=dict(color=TEKS, size=13))
    fig.update_yaxes(gridcolor="#EDF1F5", zeroline=False, linecolor=GARIS, tickfont=dict(color=TEKS, size=13),
                     title_font=dict(color=TEKS, size=13))
    return fig


def penjelasan(isi, judul_kotak="Cara membaca"):
    st.markdown(f"<div class='baca'><div class='bj'>💡 {judul_kotak}</div>{isi}</div>", unsafe_allow_html=True)


def temuan(isi):
    st.markdown(f"<div class='temuan'><b>Temuan:</b> {isi}</div>", unsafe_allow_html=True)


def subjudul(t, s=""):
    st.markdown(f"<div class='sj'>{t}</div>" + (f"<div class='ss'>{s}</div>" if s else ""), unsafe_allow_html=True)


def beda_persen(f, v, ref):
    """Selisih terhadap pembanding dalam persen; POSITIF = lebih baik, NEGATIF = lebih kurang."""
    if pd.isna(v) or pd.isna(ref) or ref == 0:
        return 0.0
    return (ref - v) / ref * 100 if kecil_baik(f) else (v - ref) / ref * 100


def grafik_kekuatan(nilai, ref, tinggi=360):
    """Diagram batang dua arah: kanan (hijau) = lebih baik dari rata-rata Tuban, kiri (merah) = lebih kurang."""
    baris = []
    for f in FITUR:
        d = beda_persen(f, nilai[f], ref[f])
        baris.append((f"{IND[f]['ikon']} {IND[f]['ukuran']}", d, kalimat(f, nilai[f]), kalimat(f, ref[f])))
    baris.sort(key=lambda b: b[1])
    warna = [HIJAU_T if d >= 10 else ("#C0392B" if d <= -10 else "#9AA5B1") for _, d, _, _ in baris]
    teks = [f"{'▲' if d >= 10 else ('▼' if d <= -10 else '●')} {fmt(abs(d))}% {'lebih baik' if d > 0 else 'lebih kurang'}"
            if abs(d) >= 1 else "● sama" for _, d, _, _ in baris]
    BATAS_TAMPIL = 100                                  # batang di atas 100% dipotong agar batang lain tetap terlihat
    batas = min(max(30, max(abs(d) for _, d, _, _ in baris) * 1.45), BATAS_TAMPIL * 1.45)
    teks = [t + (" ✂" if abs(d) > BATAS_TAMPIL else "") for t, (_, d, _, _) in zip(teks, baris)]
    fig = go.Figure(go.Bar(x=[max(min(d, BATAS_TAMPIL), -BATAS_TAMPIL) for _, d, _, _ in baris], y=[b[0] for b in baris], orientation="h", marker_color=warna,
                           text=teks, textposition="outside", textfont=dict(color=TEKS, size=13), cliponaxis=False,
                           customdata=[[b[2], b[3]] for b in baris],
                           hovertemplate="<b>%{y}</b><br>%{customdata[0]}<br>Rata-rata Tuban: %{customdata[1]}<extra></extra>"))
    fig.add_vline(x=0, line_color=TEKS, line_width=1.5)
    fig.add_vrect(x0=-10, x1=10, fillcolor="#ECEFF3", opacity=0.6, line_width=0, layer="below")
    fig.update_xaxes(range=[-batas, batas], ticksuffix="%", title="Selisih dibanding rata-rata Tuban")
    fig.add_annotation(x=-batas * 0.98, y=1.02, yref="paper", xanchor="left", showarrow=False,
                       text="◀ lebih kurang", font=dict(color="#9B1C1C", size=13))
    fig.add_annotation(x=batas * 0.98, y=1.02, yref="paper", xanchor="right", showarrow=False,
                       text="lebih baik ▶", font=dict(color=HIJAU_T, size=13))
    fig.update_layout(margin=dict(t=34))
    return gaya(fig, tinggi)


def arah_teks(f):
    return "makin <b>pendek</b> batangnya makin baik" if kecil_baik(f) else "makin <b>panjang</b> batangnya makin baik"


# Salinan cadangan (base64): dipakai hanya jika file Excel / logo tidak ditemukan. Jangan diubah.
_DATA_CADANGAN = "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"
_LOGO_CADANGAN = "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"

# =============================================================================================
# 3. DATA
# =============================================================================================
data = muat_data()
if not data_valid(data):
    st.error("Kolom data tidak lengkap. Gunakan file Data_Projek_2_Python.xlsx.")
    st.stop()
with st.spinner("Menyiapkan hasil segmentasi..."):
    rata, per_tahun = hitung(data)
TAHUN = sorted(data["Tahun"].unique())
TH0, TH1 = TAHUN[0], TAHUN[-1]
akhir = data[data["Tahun"] == TH1].set_index("Kecamatan")
rata["Penduduk"] = akhir["Jumlah_Penduduk"]
KAB = rata[FITUR].mean()                                    # rata-rata 20 kecamatan
SEG = rata.groupby("Segmen")[FITUR].mean().reindex(URUT)    # profil segmen
anggota = {s: sorted(rata.index[rata["Segmen"] == s]) for s in URUT}
pend_seg = rata.groupby("Segmen")["Penduduk"].sum().reindex(URUT)

# =============================================================================================
# 4. TAMPILAN (CSS) — semua teks berwarna gelap di atas latar terang agar selalu terbaca
# =============================================================================================
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root {{ color-scheme: light; }}
html, body, .stApp, [class*="css"] {{ font-family:'Inter','Segoe UI',Arial,sans-serif; }}
.stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {{ background:{LATAR} !important; color:{TEKS} !important; }}
header[data-testid="stHeader"] {{ background:transparent !important; }}
[data-testid="stSidebar"], [data-testid="collapsedControl"] {{ display:none; }}
.block-container {{ padding:1.4rem 2.2rem 2rem 2.2rem !important; max-width:100% !important; }}
.stApp p, .stApp li, .stApp span, .stApp label, .stApp div {{ color:{TEKS}; }}
[data-testid="stMarkdownContainer"] p {{ font-size:1rem; line-height:1.6; }}

/* Kepala halaman */
.head {{ display:flex; align-items:center; gap:18px; background:#FFFFFF; border:1px solid {GARIS}; border-radius:16px;
        padding:18px 22px; border-bottom:4px solid {NAVY}; }}
.head img {{ height:60px; }}
.head .t {{ font-size:1.45rem; font-weight:800; color:{NAVY} !important; line-height:1.25; }}
.head .s {{ font-size:.95rem; color:{ABU} !important; margin-top:4px; }}
.peta {{ background:{NAVY}; border-radius:14px; padding:16px 22px; display:flex; align-items:center;
        justify-content:space-between; gap:14px; flex-wrap:wrap; margin-top:12px; }}
.peta .pt {{ color:#FFFFFF !important; font-weight:800; font-size:1.08rem; }}
.peta .ps {{ color:#E3ECF7 !important; font-size:.93rem; margin-top:2px; }}
.peta a {{ background:#FFFFFF; color:{NAVY} !important; font-weight:800; text-decoration:none !important;
          padding:10px 20px; border-radius:10px; white-space:nowrap; font-size:.98rem; }}
.peta a:hover {{ background:#E3ECF7; }}

/* Kartu angka */
.kpi {{ background:#FFFFFF; border:1px solid {GARIS}; border-radius:14px; padding:14px 16px; height:100%; }}
.kpi .l {{ font-size:.85rem; color:{ABU} !important; font-weight:600; }}
.kpi .v {{ font-size:1.6rem; font-weight:800; color:{NAVY} !important; margin-top:2px; }}
.kpi .k {{ font-size:.83rem; color:{ABU} !important; }}

/* Tab (berlaku untuk Streamlit versi lama maupun baru) */
.stTabs [role="tablist"] {{ gap:6px; flex-wrap:nowrap !important; overflow-x:auto; border-bottom:2px solid {GARIS}; }}
.stTabs [role="tab"] {{ background:#FFFFFF !important; border:1px solid {GARIS}; border-bottom:none; border-radius:10px 10px 0 0;
        padding:10px 12px !important; height:auto; flex:1 1 0; justify-content:center; text-align:center; white-space:nowrap; }}
.stTabs [role="tab"] p {{ font-size:.98rem; font-weight:700; color:{ABU} !important; }}
.stTabs [role="tab"][aria-selected="true"] {{ background:{NAVY} !important; border-color:{NAVY}; }}
.stTabs [role="tab"][aria-selected="true"] p {{ color:#FFFFFF !important; }}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"], .stTabs .react-aria-SelectionIndicator {{ display:none !important; }}

/* Judul bagian dan kotak penjelasan */
.sj {{ font-size:1.2rem; font-weight:800; color:{NAVY} !important; margin:22px 0 2px 0; }}
.ss {{ font-size:.96rem; color:{ABU} !important; margin-bottom:10px; }}
.baca {{ background:#FFFFFF; border:1px solid {GARIS}; border-left:5px solid {NAVY}; border-radius:10px;
        padding:12px 16px; font-size:.96rem; line-height:1.65; color:{TEKS} !important; margin:6px 0 10px 0; }}
.baca .bj {{ font-weight:800; color:{NAVY} !important; margin-bottom:4px; }}
.baca b, .temuan b {{ color:{NAVY} !important; }}
.temuan {{ background:#FFF8E6; border:1px solid #F1DDA8; border-radius:10px; padding:12px 16px; font-size:.96rem;
          line-height:1.65; color:#3D3000 !important; margin:6px 0 10px 0; }}
.temuan * {{ color:#3D3000 !important; }}
.temuan b {{ color:#3D3000 !important; }}

/* Label segmen dan status */
.seg {{ display:inline-block; color:#FFFFFF !important; font-weight:700; border-radius:999px; padding:3px 12px;
       font-size:.88rem; white-space:nowrap; }}
.seg.kecil {{ font-size:.78rem; padding:2px 9px; }}
.chip {{ display:inline-block; font-weight:700; border-radius:999px; padding:2px 10px; font-size:.82rem; white-space:nowrap; }}

/* Kartu segmen */
.kartu {{ background:#FFFFFF; border:1px solid {GARIS}; border-top:6px solid var(--c); border-radius:14px; padding:16px;
         height:100%; }}
.kartu .nm {{ font-weight:800; font-size:1.05rem; color:var(--c) !important; }}
.kartu .jm {{ font-size:.88rem; color:{ABU} !important; margin:2px 0 8px 0; }}
.kartu .an {{ font-size:.92rem; line-height:1.55; margin-bottom:8px; }}
.kartu .ci {{ font-size:.9rem; line-height:1.55; border-top:1px solid #EDF1F5; padding-top:8px; }}
.kartu .sr {{ font-size:.9rem; line-height:1.5; margin-top:8px; background:var(--m); border-radius:8px; padding:8px 10px; }}
.kartu .sr b {{ color:var(--c) !important; }}

/* Tabel buatan sendiri */
.tb {{ width:100%; border-collapse:separate; border-spacing:0; background:#FFFFFF; border:1px solid {GARIS};
      border-radius:12px; overflow:hidden; font-size:.93rem; }}
.tb th {{ background:#EEF2F7; color:{NAVY} !important; font-weight:800; text-align:left; padding:10px 12px;
         border-bottom:1px solid {GARIS}; vertical-align:bottom; }}
.tb td {{ padding:9px 12px; border-bottom:1px solid #EDF1F5; vertical-align:top; color:{TEKS} !important; }}
.tb tr:last-child td {{ border-bottom:none; }}
.tb .nilai {{ font-weight:700; color:{TEKS} !important; }}
.tb .kab {{ background:#F7F9FC; }}
.wrap {{ overflow-x:auto; }}

/* Profil kecamatan */
.prof {{ background:#FFFFFF; border:1px solid {GARIS}; border-left:8px solid var(--c); border-radius:14px; padding:16px 20px; }}
.prof .a {{ font-size:1.4rem; font-weight:800; color:{NAVY} !important; }}
.prof .b {{ font-size:.97rem; line-height:1.6; margin-top:6px; }}
.ind {{ background:#FFFFFF; border:1px solid {GARIS}; border-radius:12px; padding:12px 14px; height:100%; }}
.ind .h {{ font-size:.88rem; font-weight:700; color:{ABU} !important; }}
.ind .v {{ font-size:1.2rem; font-weight:800; color:{NAVY} !important; margin:2px 0; }}
.ind .r {{ font-size:.85rem; color:{ABU} !important; margin-bottom:6px; }}

/* Widget: paksa latar putih & teks gelap (juga saat perangkat memakai mode gelap) */
[data-testid="stWidgetLabel"] p {{ color:{NAVY} !important; font-weight:700 !important; font-size:.97rem !important; }}
div[data-baseweb="select"] > div {{ background:#FFFFFF !important; border-color:#B8C4D3 !important; }}
div[data-baseweb="select"] *, div[data-baseweb="tag"] * {{ color:{TEKS} !important; }}
div[data-baseweb="tag"] {{ background:#E6EEF7 !important; }}
div[data-baseweb="popover"] *, ul[role="listbox"], ul[role="listbox"] li {{ background:#FFFFFF !important; color:{TEKS} !important; }}
ul[role="listbox"] li:hover, li[aria-selected="true"] {{ background:#E6EEF7 !important; }}
[data-testid="stSelectbox"] [role="group"], [data-testid="stMultiSelect"] [role="group"] {{ background:#FFFFFF !important;
        border:1px solid #B8C4D3 !important; }}
[data-testid="stSelectbox"] input, [data-testid="stMultiSelect"] input {{ color:{TEKS} !important; -webkit-text-fill-color:{TEKS} !important;
        background:#FFFFFF !important; }}
[data-testid="stSelectbox"] svg {{ color:{TEKS} !important; fill:{TEKS} !important; }}
[role="listbox"], .react-aria-Popover, [data-testid="stSelectboxVirtualDropdown"] {{ background:#FFFFFF !important; }}
[role="listbox"] [role="option"] {{ background:#FFFFFF !important; color:{TEKS} !important; }}
[role="listbox"] [role="option"] * {{ color:{TEKS} !important; }}
[role="listbox"] [role="option"]:hover, [role="listbox"] [role="option"][data-focused="true"],
[role="listbox"] [role="option"][aria-selected="true"] {{ background:#E6EEF7 !important; }}
[role="radiogroup"] label p, [data-testid="stRadio"] label p {{ color:{TEKS} !important; }}
[data-testid="stExpander"] {{ background:#FFFFFF; border:1px solid {GARIS}; border-radius:12px; }}
[data-testid="stExpander"] summary p {{ color:{NAVY} !important; font-weight:700; }}
.stDownloadButton button, .stButton button {{ background:{NAVY} !important; border:none !important; }}
.stDownloadButton button p, .stButton button p {{ color:#FFFFFF !important; font-weight:700; }}
.foot {{ margin-top:30px; padding-top:12px; border-top:1px solid {GARIS}; font-size:.85rem; color:{ABU} !important; }}
@media (max-width: 640px) {{ .head {{ flex-direction:column; align-items:flex-start; }} .head .t {{ font-size:1.2rem; }} }}
</style>
""", unsafe_allow_html=True)

# =============================================================================================
# 5. KEPALA HALAMAN
# =============================================================================================
st.markdown(f"""
<div class='head'><img src='{logo_src()}' alt='Logo'>
<div><div class='t'>Dashboard Ketersediaan Fasilitas Pendidikan dan Kesehatan per Kecamatan</div>
<div class='s'>Kabupaten Tuban · data {TH0}–{TH1} · sumber: Kecamatan Dalam Angka, BPS Kabupaten Tuban · Projek 2 Magang MBKM</div></div></div>
<div class='peta'><div><div class='pt'>🗺️ Peta Segmentasi Kecamatan (QGIS)</div>
<div class='ps'>Lihat letak keempat segmen di peta Kabupaten Tuban.</div></div>
<a href='{LINK_PETA}' target='_blank' rel='noopener'>Buka Peta ↗</a></div>
""", unsafe_allow_html=True)
st.write("")

kol = st.columns(4)
sekolah = akhir[["Sekolah_SD_MI", "Sekolah_SMP_MTs", "Sekolah_SMA_SMK_MA"]].sum().sum()
kpi = [("Kecamatan yang dianalisis", "20", "seluruh kecamatan di Kabupaten Tuban"),
       ("Segmen kecamatan", str(K), "kelompok dengan kondisi fasilitas mirip"),
       (f"Penduduk {TH1}", fmt(akhir["Jumlah_Penduduk"].sum()), "jiwa"),
       (f"Sekolah SD–SMA {TH1}", fmt(sekolah), "termasuk madrasah")]
for c, (l, v, s) in zip(kol, kpi):
    c.markdown(f"<div class='kpi'><div class='l'>{l}</div><div class='v'>{v}</div><div class='k'>{s}</div></div>",
               unsafe_allow_html=True)

tab = st.tabs(["🏠 Ringkasan", "🧩 Empat Segmen", "📊 Bandingkan Segmen", "🏆 Peringkat Kecamatan",
               "📍 Profil Kecamatan", "📈 Perkembangan 2022–2025", "ℹ️ Tentang Data"])

# =============================================================================================
# TAB 1. RINGKASAN
# =============================================================================================
with tab[0]:
    subjudul("Apa isi dashboard ini?",
             "Dashboard ini menjawab satu pertanyaan: kecamatan mana yang fasilitas pendidikan dan kesehatannya sudah cukup, dan mana yang masih kurang?")
    penjelasan(
        "Fasilitas tidak dibandingkan dari <b>jumlahnya</b> saja, karena kecamatan yang penduduknya banyak pasti punya sekolah lebih banyak. "
        "Karena itu setiap fasilitas dihitung <b>dibanding jumlah penggunanya</b>, misalnya "
        "<b>“1 SD melayani berapa penduduk”</b> atau <b>“1 guru mengajar berapa murid”</b>. "
        "Lalu 20 kecamatan dikelompokkan menjadi <b>4 segmen</b> berdasarkan kemiripan kondisi tujuh indikator, memakai metode K-Means.",
        "Ide dasarnya")

    subjudul("Tujuh indikator yang dipakai", "Semua indikator memakai rata-rata 4 tahun agar tidak terpengaruh perubahan satu tahun saja.")
    baris = ""
    for f, v in IND.items():
        lebih = "lebih <b>sedikit</b> = lebih baik" if kecil_baik(f) else "lebih <b>banyak</b> = lebih baik"
        baris += (f"<tr><td>{v['ikon']} <b>{v['nama']}</b><br><span style='color:{ABU}'>{v['kelompok']}</span></td>"
                  f"<td>{v['ukuran']}</td><td>{lebih}</td>"
                  f"<td class='nilai'>{kalimat(f, ubah(f, KAB[f]))}</td></tr>")
    st.markdown(f"<div class='wrap'><table class='tb'><tr><th>Indikator</th><th>Diukur sebagai</th><th>Arah yang baik</th>"
                f"<th>Rata-rata 20 kecamatan</th></tr>{baris}</table></div>", unsafe_allow_html=True)
    _sma = fmt(bulat("Rasio_Sekolah_SMA", ubah("Rasio_Sekolah_SMA", KAB["Rasio_Sekolah_SMA"])))
    penjelasan(f"Baca kolom paling kanan seperti kalimat biasa. Contohnya, “1 SMA/SMK/MA melayani ± {_sma} penduduk” artinya "
               f"secara rata-rata satu SMA sederajat harus menampung kebutuhan sekitar {_sma} orang. "
               "Untuk sekolah, puskesmas, dan guru, angka yang lebih <b>kecil</b> berarti fasilitas lebih longgar. "
               "Untuk apotek, angka yang lebih <b>besar</b> berarti apotek lebih merata.")

    subjudul("Hasil utama dalam satu tabel", "Ringkasan empat segmen: siapa anggotanya, apa cirinya, dan apa yang perlu dilakukan.")
    baris = ""
    for s in URUT:
        baris += (f"<tr><td>{seg_label(s)}<br><span style='color:{ABU};font-size:.85rem'>{len(anggota[s])} kecamatan · "
                  f"{fmt(pend_seg[s] / pend_seg.sum() * 100, 1)}% penduduk</span></td>"
                  f"<td>{', '.join(anggota[s])}</td><td>{CIRI[s]}</td><td><b>{SARAN[s]}</b></td></tr>")
    st.markdown(f"<div class='wrap'><table class='tb'><tr><th style='width:19%'>Segmen</th><th style='width:22%'>Kecamatan</th>"
                f"<th>Ciri utama</th><th style='width:23%'>Rekomendasi</th></tr>{baris}</table></div>", unsafe_allow_html=True)
    temuan(f"Tidak ada segmen yang unggul di semua hal. Setiap segmen punya kekurangan yang berbeda, sehingga kebijakannya pun perlu berbeda. "
           f"Segmen <b>{PP}</b> paling perlu perhatian karena menampung <b>{fmt(pend_seg[PP] / pend_seg.sum() * 100, 1)}%</b> penduduk Tuban.")

# =============================================================================================
# TAB 2. EMPAT SEGMEN
# =============================================================================================
with tab[1]:
    subjudul("Kenali empat segmen", "Setiap kartu menjelaskan satu segmen: anggota, ciri, dan rekomendasinya.")
    kol = st.columns(4)
    for c, s in zip(kol, URUT):
        c.markdown(f"""<div class='kartu' style='--c:{WARNA[s]};--m:{MUDA[s]}'>
            <div class='nm'>{s}</div>
            <div class='jm'>{len(anggota[s])} kecamatan · {fmt(pend_seg[s])} penduduk</div>
            <div class='an'>{', '.join(anggota[s])}</div>
            <div class='ci'>{CIRI[s]}</div>
            <div class='sr'><b>Rekomendasi:</b> {SARAN[s]}</div></div>""", unsafe_allow_html=True)

    subjudul("Seberapa besar setiap segmen?",
             "Kiri: berapa kecamatan di setiap segmen. Kanan: berapa penduduk " + str(TH1) + " yang tinggal di setiap segmen.")
    kiri, kanan = st.columns([2, 3])
    with kiri:
        jml = pd.Series({s: len(anggota[s]) for s in URUT})
        fig = go.Figure(go.Pie(labels=jml.index, values=jml.values, hole=0.55, sort=False, direction="clockwise",
                               marker=dict(colors=[WARNA[s] for s in jml.index], line=dict(color="#FFFFFF", width=2)),
                               texttemplate="%{value}", textfont=dict(color="#FFFFFF", size=16),
                               hovertemplate="<b>%{label}</b><br>%{value} kecamatan<extra></extra>"))
        fig.add_annotation(text="<b>20</b><br>kecamatan", showarrow=False, font=dict(size=16, color=NAVY))
        fig.update_layout(legend=dict(orientation="v", y=0.5, x=1.0, yanchor="middle"))
        grafik(gaya(fig, 300))
    p = pend_seg.reindex(URUT[::-1])
    fig = go.Figure(go.Bar(x=p.values, y=p.index, orientation="h", marker_color=[WARNA[s] for s in p.index],
                           text=[f"{fmt(v)} jiwa ({fmt(v / p.sum() * 100, 1)}%) · {len(anggota[s])} kecamatan" for s, v in p.items()],
                           textposition="outside", textfont=dict(color=TEKS, size=14), cliponaxis=False,
                           hovertemplate="<b>%{y}</b><br>%{x:,.0f} jiwa<extra></extra>"))
    fig.update_xaxes(title="Jumlah penduduk (jiwa)", range=[0, p.max() * 1.6])
    with kanan:
        grafik(gaya(fig, 300))
    penjelasan("<b>Diagram lingkaran (kiri):</b> setiap potongan adalah satu segmen, angkanya jumlah kecamatan dari total 20. "
               "<b>Diagram batang (kanan):</b> makin panjang batangnya, makin banyak penduduk yang tinggal di segmen itu. "
               "Angka di ujung batang menunjukkan jumlah penduduk dan persentasenya dari seluruh Tuban. "
               "Bandingkan keduanya: segmen dengan sedikit kecamatan belum tentu penduduknya sedikit.")
    temuan(f"<b>{PP}</b> berisi {len(anggota[PP])} kecamatan dan menampung hampir separuh penduduk Tuban. "
           f"<b>{PK}</b> hanya 2 kecamatan, tetapi sangat padat.")

    subjudul("Seberapa yakin pengelompokan ini?",
             "Pengelompokan diulang dengan data yang berbeda-beda (hanya satu tahun, tanpa tahun tertentu, tanpa Kecamatan Tuban).")
    st_ = rata[["Segmen", "Stabil_%"]].copy()
    st_["Status"] = np.where(st_["Stabil_%"] >= 80, "Anggota inti (stabil)", "Di perbatasan dua segmen")
    n_inti = int((st_["Stabil_%"] >= 80).sum())
    kol = st.columns([1, 2])
    kol[0].markdown(f"<div class='kpi'><div class='l'>Kecamatan yang segmennya stabil</div>"
                    f"<div class='v'>{n_inti} dari 20</div><div class='k'>tetap di segmen yang sama pada ≥ 80% percobaan</div></div>",
                    unsafe_allow_html=True)
    with kol[1]:
        penjelasan("Untuk menguji apakah hasil ini bisa dipercaya, pengelompokan diulang beberapa kali dengan data yang sedikit berbeda. "
                   "Kecamatan yang hampir selalu masuk segmen yang sama disebut <b>anggota inti</b>. "
                   "Makin banyak anggota inti, makin bisa dipercaya hasilnya untuk dasar kebijakan.", "Apa artinya?")

# =============================================================================================
# TAB 3. BANDINGKAN SEGMEN
# =============================================================================================
with tab[2]:
    subjudul("Bandingkan keempat segmen untuk satu indikator", "Pilih indikator yang ingin dilihat.")
    pil = st.selectbox("Pilih indikator", list(PILIHAN), key="pil_seg")
    f = PILIHAN[pil]
    nilai = SEG[f].map(lambda x: ubah(f, x))
    kab = ubah(f, KAB[f])
    urut = nilai.sort_values(ascending=not kecil_baik(f))            # urutan dari paling kurang ke paling baik
    gab = pd.concat([urut, pd.Series({"Rata-rata Tuban": kab})]).sort_values(ascending=not kecil_baik(f))
    fig = go.Figure(go.Bar(x=gab.values, y=gab.index, orientation="h",
                           marker_color=[WARNA.get(s, "#9AA5B1") for s in gab.index],
                           marker_pattern_shape=["/" if s == "Rata-rata Tuban" else "" for s in gab.index],
                           text=[teks_nilai(f, v) for v in gab.values], textposition="outside",
                           textfont=dict(color=TEKS, size=14), cliponaxis=False,
                           customdata=[kalimat(f, v) for v in gab.values],
                           hovertemplate="<b>%{y}</b><br>%{customdata}<extra></extra>"))
    fig.update_xaxes(title=IND[f]["ukuran"] + (" (%)" if IND[f]["jenis"] == "persen" else ""),
                     range=[0, gab.max() * 1.3])
    grafik(gaya(fig, 340))
    terbaik, terburuk = urut.index[-1], urut.index[0]
    penjelasan(f"Setiap batang berwarna adalah rata-rata satu segmen, dan batang abu-abu bergaris adalah rata-rata seluruh Tuban "
               f"sebagai pembanding. Untuk indikator ini, {arah_teks(f)}. Batang sudah diurutkan: "
               f"segmen yang paling baik ada di <b>paling atas</b>, yang paling kurang di <b>paling bawah</b>.")
    temuan(f"Paling baik: <b>{terbaik}</b> ({kalimat(f, nilai[terbaik])}). "
           f"Paling kurang: <b>{terburuk}</b> ({kalimat(f, nilai[terburuk])}).")

    subjudul("Rapor segmen: semua indikator sekaligus",
             "Setiap sel berisi nilai segmen dan perbandingannya dengan rata-rata Tuban.")
    kepala = "".join(f"<th>{seg_label(s, True)}</th>" for s in URUT)
    baris = ""
    for f2, v in IND.items():
        k2 = ubah(f2, KAB[f2])
        sel = "".join(f"<td><div class='nilai'>{teks_nilai(f2, ubah(f2, SEG.loc[s, f2]))}</div>"
                      f"{chip(status(f2, ubah(f2, SEG.loc[s, f2]), k2))}</td>" for s in URUT)
        baris += (f"<tr><td><b>{v['ikon']} {v['ukuran']}</b></td>{sel}"
                  f"<td class='kab'><div class='nilai'>{teks_nilai(f2, k2)}</div></td></tr>")
    st.markdown(f"<div class='wrap'><table class='tb'><tr><th>Indikator</th>{kepala}<th>Rata-rata Tuban</th></tr>"
                f"{baris}</table></div>", unsafe_allow_html=True)
    penjelasan("Baca per baris. Label <span class='chip' style='color:#1B5E20;background:#E3F4E8'>▲ Lebih baik</span> berarti "
               "segmen itu lebih baik dari rata-rata Tuban, <span class='chip' style='color:#9B1C1C;background:#FDE8E8'>▼ Lebih kurang</span> "
               "berarti lebih buruk, dan <span class='chip' style='color:#37474F;background:#ECEFF3'>● Setara rata-rata</span> berarti "
               "selisihnya kurang dari 10%. Contohnya, nilai “" + teks_nilai("Rasio_Sekolah_SMA", ubah("Rasio_Sekolah_SMA", SEG.loc[PM, "Rasio_Sekolah_SMA"]))
               + "” pada segmen " + PM + " di baris SMA artinya 1 SMA melayani sekitar sebanyak itu.")

    subjudul("Kekuatan dan kelemahan setiap segmen",
             "Pilih satu segmen untuk melihat indikator mana yang lebih baik dan mana yang lebih kurang dari rata-rata Tuban.")
    pil_s = st.radio("Pilih segmen", URUT, horizontal=True, key="seg_kk")
    grafik(grafik_kekuatan({f: ubah(f, SEG.loc[pil_s, f]) for f in FITUR}, {f: ubah(f, KAB[f]) for f in FITUR}))
    dd = {f: beda_persen(f, ubah(f, SEG.loc[pil_s, f]), ubah(f, KAB[f])) for f in FITUR}
    fb, fk = max(dd, key=dd.get), min(dd, key=dd.get)
    penjelasan("Batang ke <b>kanan (hijau)</b> berarti segmen ini <b>lebih baik</b> dari rata-rata Tuban, batang ke <b>kiri (merah)</b> "
               "berarti <b>lebih kurang</b>. Panjang batang menunjukkan seberapa jauh bedanya dalam persen. "
               "Daerah abu-abu di tengah (−10% sampai +10%) berarti hampir sama dengan rata-rata. "
               "Tanda ✂ berarti batangnya dipotong di 100% agar batang lain tetap terlihat; angkanya tetap nilai sebenarnya.")
    temuan(f"Kekuatan terbesar <b>{pil_s}</b>: {IND[fb]['ukuran']} ({fmt(abs(dd[fb]))}% lebih baik). "
           f"Kelemahan terbesar: {IND[fk]['ukuran']} ({fmt(abs(dd[fk]))}% lebih kurang). "
           f"Rekomendasi: {SARAN[pil_s]}")

# =============================================================================================
# TAB 4. PERINGKAT KECAMATAN
# =============================================================================================
with tab[3]:
    subjudul("Peringkat 20 kecamatan", "Pilih indikator, lalu lihat kecamatan mana yang paling baik dan paling kurang.")
    k1, k2 = st.columns([2, 1])
    pil = k1.selectbox("Pilih indikator", list(PILIHAN), key="pil_rank")
    sorot = k2.selectbox("Sorot kecamatan (opsional)", ["–"] + sorted(rata.index), key="sorot")
    f = PILIHAN[pil]
    v = rata[f].map(lambda x: ubah(f, x))
    kab = ubah(f, KAB[f])
    v = pd.concat([v, pd.Series({"Rata-rata Tuban": kab})]).sort_values(ascending=kecil_baik(f))   # terbaik di atas
    seg_v = [rata.loc[k, "Segmen"] if k in rata.index else "Pembanding" for k in v.index]
    warna = [WARNA.get(sg, "#9AA5B1") for sg in seg_v]
    garis_bar = [3 if k == sorot else 0 for k in v.index]
    label_y = [f"<b>▶ {k}</b>" if k == sorot else k for k in v.index]
    fig = go.Figure(go.Bar(x=v.values, y=label_y, orientation="h",
                           marker=dict(color=warna, line=dict(color="#111111", width=garis_bar),
                                       pattern_shape=["/" if k == "Rata-rata Tuban" else "" for k in v.index]),
                           text=[teks_nilai(f, x) for x in v.values], textposition="outside",
                           textfont=dict(color=TEKS, size=13), cliponaxis=False,
                           customdata=np.stack([seg_v, [kalimat(f, x) for x in v.values]], axis=1),
                           hovertemplate="<b>%{y}</b><br>%{customdata[1]}<br>Segmen: %{customdata[0]}<extra></extra>",
                           showlegend=False))
    for sg in URUT:                                              # legenda warna segmen
        fig.add_trace(go.Bar(x=[None], y=[None], marker_color=WARNA[sg], name=sg, orientation="h"))
    fig.update_xaxes(title=IND[f]["ukuran"] + (" (%)" if IND[f]["jenis"] == "persen" else ""), range=[0, v.max() * 1.25])
    fig.update_yaxes(tickfont=dict(size=13), autorange="reversed")
    fig.update_layout(margin=dict(t=50), legend=dict(y=1.02))
    grafik(gaya(fig, 680))
    v = v.drop("Rata-rata Tuban")
    penjelasan(f"Setiap batang adalah satu kecamatan, warnanya menunjukkan segmennya, dan batang abu-abu bergaris adalah rata-rata Tuban. "
               f"Untuk indikator ini, {arah_teks(f)}. Kecamatan sudah diurutkan: yang <b>paling atas</b> paling baik, "
               f"yang <b>paling bawah</b> paling kurang. Kecamatan di atas batang abu-abu berarti lebih baik dari rata-rata Tuban. "
               f"Arahkan kursor ke batang untuk melihat kalimat lengkapnya.")
    lebih_baik = int(sum((x < kab) if kecil_baik(f) else (x > kab) for x in v.values))
    _temuan_rank = (v, f, kab, lebih_baik)
    temuan(f"Paling baik: <b>{v.index[0]}</b> ({kalimat(f, v.iloc[0])}). "
           f"Paling kurang: <b>{v.index[-1]}</b> ({kalimat(f, v.iloc[-1])}). "
           f"Ada <b>{lebih_baik} dari 20</b> kecamatan yang lebih baik dari rata-rata Tuban.")

    subjudul("Peta posisi: pendidikan dibanding kesehatan",
             "Lihat dua indikator sekaligus untuk menemukan kecamatan yang kurang di keduanya.")
    k1, k2 = st.columns(2)
    pil_pend = [k for k, f in PILIHAN.items() if IND[f]["kelompok"] == "Pendidikan"]
    pil_kes = [k for k, f in PILIHAN.items() if IND[f]["kelompok"] == "Kesehatan"]
    px_ = k1.selectbox("Indikator pendidikan (sumbu mendatar)", pil_pend, index=1, key="sc_x")
    py_ = k2.selectbox("Indikator kesehatan (sumbu tegak)", pil_kes, index=0, key="sc_y")
    fx, fy = PILIHAN[px_], PILIHAN[py_]
    xs = rata[fx].map(lambda x: ubah(fx, x)); ys = rata[fy].map(lambda x: ubah(fy, x))
    kx, ky = ubah(fx, KAB[fx]), ubah(fy, KAB[fy])
    rx = [xs.min() * 0.85, xs.max() * 1.1]; ry = [ys.min() * 0.85, ys.max() * 1.1]
    # Penempatan nama kecamatan: dicari posisi yang tidak menabrak nama lain / titik lain, lalu dihubungkan garis tipis
    W, H = 1500, 430
    def ke_layar(xv, yv):
        u = (xv - rx[0]) / (rx[1] - rx[0]); v = (yv - ry[0]) / (ry[1] - ry[0])
        u = 1 - u if kecil_baik(fx) else u; v = 1 - v if kecil_baik(fy) else v
        return u * W, (1 - v) * H
    layar = {k: ke_layar(xs[k], ys[k]) for k in rata.index}
    kotak, label = [], {}
    for k in sorted(rata.index, key=lambda k: layar[k][0]):
        sx, sy = layar[k]; lw, lh = 7.2 * len(k) + 8, 17
        terbaik_ = None
        for d in (20, 34, 50, 68, 88):
            for sudut in (90, 270, 30, 150, 330, 210, 0, 180):
                ax_ = d * np.cos(np.radians(sudut)) * (1.6 if sudut in (0, 180, 30, 150, 330, 210) else 1)
                ay_ = -d * np.sin(np.radians(sudut))
                cx, cy = sx + ax_, sy + ay_
                n = sum(1 for (x0, y0, w0, h0) in kotak if abs(cx - x0) < (lw + w0) / 2 and abs(cy - y0) < (lh + h0) / 2)
                n += sum(1 for k2, (px2, py2) in layar.items() if abs(cx - px2) < lw / 2 + 8 and abs(cy - py2) < lh / 2 + 8)
                if terbaik_ is None or n < terbaik_[0]:
                    terbaik_ = (n, ax_, ay_, cx, cy)
                if n == 0:
                    break
            if terbaik_[0] == 0:
                break
        _, ax_, ay_, cx, cy = terbaik_
        kotak.append((cx, cy, lw, lh)); label[k] = (ax_, ay_)
    fig = go.Figure()
    for sg in URUT:
        m = rata["Segmen"] == sg
        fig.add_trace(go.Scatter(x=xs[m], y=ys[m], mode="markers", name=sg, text=list(rata.index[m]),
                                 marker=dict(size=15, color=WARNA[sg], line=dict(color="#FFFFFF", width=1.5)),
                                 customdata=[[kalimat(fx, a), kalimat(fy, b)] for a, b in zip(xs[m], ys[m])],
                                 hovertemplate="<b>%{text}</b><br>%{customdata[0]}<br>%{customdata[1]}<extra>" + sg + "</extra>"))
    for k, (ax_, ay_) in label.items():
        fig.add_annotation(x=xs[k], y=ys[k], ax=ax_, ay=ay_, text=k, showarrow=True, arrowhead=0, arrowwidth=1,
                           arrowcolor="#9AA5B1", standoff=7, font=dict(color=TEKS, size=12), bgcolor="rgba(255,255,255,0.8)")
    fig.add_vline(x=kx, line_dash="dash", line_color="#6B7785"); fig.add_hline(y=ky, line_dash="dash", line_color="#6B7785")
    fig.update_xaxes(title=IND[fx]["ukuran"] + (" (%)" if IND[fx]["jenis"] == "persen" else ""),
                     range=rx[::-1] if kecil_baik(fx) else rx)
    fig.update_yaxes(title=IND[fy]["ukuran"] + (" (%)" if IND[fy]["jenis"] == "persen" else ""),
                     range=ry[::-1] if kecil_baik(fy) else ry)
    sudut = [(0.99, 0.99, "right", "top", "✅ Keduanya lebih baik", HIJAU_T),
             (0.01, 0.01, "left", "bottom", "⚠️ Keduanya lebih kurang", "#9B1C1C"),
             (0.01, 0.99, "left", "top", "Kurang " + IND[fx]["nama"], "#7A4A00"),
             (0.99, 0.01, "right", "bottom", "Kurang " + IND[fy]["nama"], "#7A4A00")]
    for xx, yy, xa, ya, t, c in sudut:
        fig.add_annotation(x=xx, y=yy, xref="paper", yref="paper", xanchor=xa, yanchor=ya, showarrow=False, text=f"<b>{t}</b>",
                           font=dict(color=c, size=13), bgcolor="rgba(255,255,255,0.85)")
    grafik(gaya(fig, 560))
    kanan_atas = [k for k in rata.index if beda_persen(fx, xs[k], kx) > 0 and beda_persen(fy, ys[k], ky) > 0]
    kiri_bawah = [k for k in rata.index if beda_persen(fx, xs[k], kx) < 0 and beda_persen(fy, ys[k], ky) < 0]
    penjelasan("Setiap titik adalah satu kecamatan, warnanya menunjukkan segmen. Garis putus-putus adalah rata-rata Tuban dan membagi "
               "grafik menjadi empat bagian. Sumbu sudah diatur sehingga <b>makin ke kanan dan makin ke atas berarti makin baik</b>. "
               "Pojok kanan atas = kedua fasilitas lebih baik dari rata-rata; pojok kiri bawah = keduanya lebih kurang dan paling perlu diprioritaskan.")
    temuan(f"Kecamatan yang <b>kurang di kedua indikator</b>: {', '.join(kiri_bawah) or 'tidak ada'}. "
           f"Kecamatan yang <b>lebih baik di keduanya</b>: {', '.join(kanan_atas) or 'tidak ada'}.")

# =============================================================================================
# TAB 5. PROFIL KECAMATAN
# =============================================================================================
with tab[4]:
    subjudul("Profil satu kecamatan", "Pilih kecamatan untuk melihat kondisi fasilitasnya dibanding rata-rata Tuban.")
    kec = st.selectbox("Pilih kecamatan", sorted(rata.index), index=sorted(rata.index).index("Kenduruan")
                       if "Kenduruan" in rata.index else 0, key="kec")
    s = rata.loc[kec, "Segmen"]
    kuat = [f for f in FITUR if status(f, ubah(f, rata.loc[kec, f]), ubah(f, KAB[f])) == "Lebih baik"]
    lemah = [f for f in FITUR if status(f, ubah(f, rata.loc[kec, f]), ubah(f, KAB[f])) == "Lebih kurang"]
    st.markdown(f"""<div class='prof' style='--c:{WARNA[s]}'>
        <div class='a'>Kecamatan {kec}</div>
        <div class='b'>Termasuk segmen {seg_label(s)} · penduduk {TH1}: <b>{fmt(rata.loc[kec, 'Penduduk'])} jiwa</b> ·
        segmen stabil di <b>{fmt(rata.loc[kec, 'Stabil_%'])}%</b> percobaan</div>
        <div class='b'><b>Kelebihan:</b> {', '.join(IND[f]['nama'] for f in kuat) or 'tidak ada yang menonjol'}.
        <b>Kekurangan:</b> {', '.join(IND[f]['nama'] for f in lemah) or 'tidak ada yang menonjol'}.</div>
        <div class='b'><b>Rekomendasi segmen:</b> {SARAN[s]}</div></div>""", unsafe_allow_html=True)
    st.write("")
    kol = st.columns(4)
    for i, f in enumerate(FITUR):
        nk, nkab = ubah(f, rata.loc[kec, f]), ubah(f, KAB[f])
        kol[i % 4].markdown(f"""<div class='ind'><div class='h'>{IND[f]['ikon']} {IND[f]['ukuran']}</div>
            <div class='v'>{teks_nilai(f, nk)}</div>
            <div class='r'>Rata-rata Tuban: {teks_nilai(f, nkab)}</div>{chip(status(f, nk, nkab))}</div>""",
                            unsafe_allow_html=True)
        if i == 3:
            st.write("")
            kol = st.columns(4)
    penjelasan(f"Setiap kotak berisi satu indikator untuk Kecamatan {kec} (rata-rata {TH0}–{TH1}). "
               "Angka besar adalah nilai kecamatan ini, di bawahnya rata-rata Tuban sebagai pembanding. "
               "Label hijau ▲ berarti lebih baik dari rata-rata, merah ▼ berarti lebih kurang, abu-abu ● berarti hampir sama (selisih < 10%).")

    subjudul(f"Kekuatan dan kelemahan Kecamatan {kec}", "Seberapa jauh setiap indikator berbeda dari rata-rata Tuban.")
    grafik(grafik_kekuatan({f: ubah(f, rata.loc[kec, f]) for f in FITUR}, {f: ubah(f, KAB[f]) for f in FITUR}))
    penjelasan("Batang ke <b>kanan (hijau)</b> = lebih baik dari rata-rata Tuban, ke <b>kiri (merah)</b> = lebih kurang. "
               "Makin panjang batangnya, makin jauh bedanya. Indikator paling bawah adalah kelemahan terbesar kecamatan ini, "
               "yang paling atas adalah kekuatan terbesarnya. Tanda ✂ berarti batang dipotong di 100%, angkanya tetap nilai sebenarnya.")

    subjudul(f"Perubahan di Kecamatan {kec} dari tahun ke tahun")
    pil = st.selectbox("Pilih indikator", list(PILIHAN), key="pil_prof")
    f = PILIHAN[pil]
    dk = data[data["Kecamatan"] == kec].set_index("Tahun")[f].map(lambda x: ubah(f, x))
    dkab = data.groupby("Tahun")[f].mean().map(lambda x: ubah(f, x))
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=dkab.index, y=dkab.values, name="Rata-rata Tuban", mode="lines+markers",
                             line=dict(color="#8A96A5", dash="dash", width=2), marker=dict(size=8),
                             hovertemplate="%{x}: %{y:,.1f}<extra>Rata-rata Tuban</extra>"))
    fig.add_trace(go.Scatter(x=dk.index, y=dk.values, name=f"Kecamatan {kec}", mode="lines+markers+text",
                             line=dict(color=WARNA[s], width=3), marker=dict(size=10),
                             text=[teks_nilai(f, x) for x in dk.values], textposition="top center",
                             textfont=dict(color=TEKS, size=13), hovertemplate="%{x}: %{y:,.1f}<extra>" + kec + "</extra>"))
    lo, hi = min(dk.min(), dkab.min()), max(dk.max(), dkab.max())
    fig.update_yaxes(title=IND[f]["ukuran"], range=[0, hi * 1.25])
    fig.update_xaxes(tickmode="array", tickvals=TAHUN, title="")
    grafik(gaya(fig, 340))
    penjelasan(f"Garis berwarna adalah Kecamatan {kec}, garis putus-putus abu-abu adalah rata-rata Tuban. "
               f"Untuk indikator ini, {('garis yang <b>turun</b> berarti membaik' if kecil_baik(f) else 'garis yang <b>naik</b> berarti membaik')}. "
               "Sumbu tegak dimulai dari nol agar perubahan kecil tidak terlihat berlebihan."
               + (" Catatan: data puskesmas dan apotek tahun 2022 disamakan dengan 2023." if f in ("Rasio_Puskesmas", "Persen_Desa_Apotek") else ""))

# =============================================================================================
# TAB 6. PERKEMBANGAN
# =============================================================================================
with tab[5]:
    subjudul(f"Bagaimana kondisi Tuban berubah dari {TH0} ke {TH1}?",
             "Rata-rata 20 kecamatan setiap tahun, dalam ukuran yang mudah dibayangkan.")
    th_rata = data.groupby("Tahun")[FITUR].mean()
    baris = ""
    for f, v in IND.items():
        a, b = ubah(f, th_rata.loc[TH0, f]), ubah(f, th_rata.loc[TH1, f])
        ubah_pct = (b / a - 1) * 100 if a else 0
        if abs(ubah_pct) < 3:
            ket = "Relatif tetap"
        else:
            ket = "Membaik" if ((ubah_pct < 0) if kecil_baik(f) else (ubah_pct > 0)) else "Memburuk"
        t, bg = {"Membaik": (HIJAU_T, HIJAU_B), "Memburuk": (MERAH_T, MERAH_B), "Relatif tetap": (NETRAL_T, NETRAL_B)}[ket]
        sel = "".join(f"<td class='nilai'>{teks_nilai(f, ubah(f, th_rata.loc[t2, f]))}</td>" for t2 in TAHUN)
        baris += (f"<tr><td><b>{v['ikon']} {v['ukuran']}</b></td>{sel}"
                  f"<td><span class='chip' style='color:{t};background:{bg}'>{ket}</span>"
                  f"<div style='font-size:.82rem;color:{ABU};margin-top:3px'>{'naik' if ubah_pct >= 0 else 'turun'} {fmt(abs(ubah_pct), 1)}%</div></td></tr>")
    kepala = "".join(f"<th>{t}</th>" for t in TAHUN)
    st.markdown(f"<div class='wrap'><table class='tb'><tr><th>Indikator</th>{kepala}<th>{TH0} → {TH1}</th></tr>{baris}</table></div>",
                unsafe_allow_html=True)
    penjelasan("Baca per baris dari kiri ke kanan. Kolom terakhir merangkum arah perubahannya: <b>Membaik</b>, <b>Memburuk</b>, "
               "atau <b>Relatif tetap</b> (berubah kurang dari 3%), beserta besar perubahannya dalam persen. Ingat: untuk penduduk per fasilitas dan murid per guru, angka yang <b>naik</b> berarti <b>memburuk</b>. Data puskesmas dan apotek tahun 2022 disamakan dengan 2023, "
               "karena data 2022 dicatat dengan cara berbeda.")
    subjudul(f"Apa yang bertambah dan berkurang dari {TH0} ke {TH1}?",
             "Perubahan jumlah (bukan rasio) untuk seluruh Kabupaten Tuban, dalam persen.")
    tot = data.groupby("Tahun")[["Jumlah_Penduduk", "Sekolah_SD_MI", "Sekolah_SMP_MTs", "Sekolah_SMA_SMK_MA", "Guru_SD_MI",
                                 "Guru_SMP_MTs", "Guru_SMA_SMK_MA", "Murid_SD_MI", "Murid_SMP_MTs", "Murid_SMA_SMK_MA",
                                 "Desa_Ada_Puskesmas", "Desa_Ada_Apotek"]].sum()
    nm = {"Jumlah_Penduduk": "👥 Penduduk", "Sekolah_SD_MI": "🏫 Sekolah SD/MI", "Sekolah_SMP_MTs": "🏫 Sekolah SMP/MTs",
          "Sekolah_SMA_SMK_MA": "🎓 Sekolah SMA/SMK/MA", "Guru_SD_MI": "👩‍🏫 Guru SD/MI", "Guru_SMP_MTs": "👨‍🏫 Guru SMP/MTs",
          "Guru_SMA_SMK_MA": "🧑‍🏫 Guru SMA/SMK/MA", "Murid_SD_MI": "🎒 Murid SD/MI", "Murid_SMP_MTs": "🎒 Murid SMP/MTs",
          "Murid_SMA_SMK_MA": "🎒 Murid SMA/SMK/MA", "Desa_Ada_Puskesmas": "🏥 Puskesmas", "Desa_Ada_Apotek": "💊 Desa ber-apotek"}
    pc = ((tot.loc[TH1] / tot.loc[TH0] - 1) * 100).rename(index=nm)
    urutan_nm = list(nm.values())[::-1]
    pc = pc.reindex(urutan_nm)
    fig = go.Figure(go.Bar(x=pc.values, y=pc.index, orientation="h",
                           marker_color=["#1F4E79" if x >= 0 else "#C05A00" for x in pc.values],
                           text=[f"{'+' if x >= 0 else '−'}{fmt(abs(x), 1)}%  ({fmt(tot.loc[TH0, k])} → {fmt(tot.loc[TH1, k])})"
                                 for k, x in zip([k for k in nm][::-1], pc.values)],
                           textposition="outside", textfont=dict(color=TEKS, size=13), cliponaxis=False,
                           hovertemplate="<b>%{y}</b><br>%{x:.1f}%<extra></extra>"))
    fig.add_vline(x=0, line_color=TEKS, line_width=1.5)
    bts = max(abs(pc.min()), abs(pc.max())) * 1.9
    fig.update_xaxes(range=[-bts, bts], ticksuffix="%", title=f"Perubahan jumlah {TH0} → {TH1}")
    grafik(gaya(fig, 480))
    penjelasan(f"Batang ke <b>kanan (biru)</b> berarti jumlahnya <b>bertambah</b>, ke <b>kiri (oranye)</b> berarti <b>berkurang</b>. "
               f"Angka dalam kurung adalah jumlah tahun {TH0} dan {TH1}. Grafik ini menjelaskan <i>mengapa</i> beban guru naik: "
               "bandingkan perubahan jumlah guru dengan perubahan jumlah murid di jenjang yang sama.")
    temuan("Jumlah sekolah dan puskesmas dibanding penduduk relatif tetap, dan apotek makin menyebar. "
           "Yang perlu diwaspadai adalah <b>beban guru</b> yang naik di semua jenjang, terutama pada tahun 2025, karena jumlah guru berkurang.")

    subjudul("Segmen setiap kecamatan jika dihitung per tahun",
             "Segmen utama dihitung dari rata-rata 4 tahun. Tabel ini menunjukkan segmennya jika memakai data satu tahun saja.")
    pt = per_tahun.pivot(index="Kecamatan", columns="Tahun", values="Segmen")
    urutan = rata.assign(_u=rata["Segmen"].map({s: i for i, s in enumerate(URUT)})).sort_values(["_u", "Penduduk"],
                                                                                              ascending=[True, False]).index
    pt = pt.loc[urutan]
    SAMA = f"<span style='color:{ABU}'>sama</span>"
    baris = ""
    for k, r in pt.iterrows():
        utama = rata.loc[k, "Segmen"]
        sel = "".join("<td>" + (seg_label(r[t], True) if r[t] != utama else SAMA) + "</td>" for t in TAHUN)
        baris += f"<tr><td><b>{k}</b></td><td>{seg_label(utama, True)}</td>{sel}</tr>"
    kepala = "".join(f"<th>{t}</th>" for t in TAHUN)
    st.markdown(f"<div class='wrap'><table class='tb'><tr><th>Kecamatan</th><th>Segmen utama</th>{kepala}</tr>{baris}</table></div>",
                unsafe_allow_html=True)
    n_pindah = int((pt.ne(rata.loc[pt.index, "Segmen"], axis=0)).any(axis=1).sum())
    penjelasan("Tulisan “sama” berarti pada tahun itu kecamatan berada di segmen utamanya. "
               "Jika muncul label segmen lain, berarti pada tahun itu kondisinya lebih mirip segmen tersebut.")
    temuan(f"{20 - n_pindah} dari 20 kecamatan selalu berada di segmen yang sama setiap tahun. "
           f"Perpindahan yang ada umumnya terjadi pada tahun {TH1}, saat jumlah guru berkurang.")

# =============================================================================================
# TAB 7. TENTANG DATA
# =============================================================================================
with tab[6]:
    subjudul("Langkah analisis (versi singkat)")
    st.markdown(f"""<div class='wrap'><table class='tb'>
    <tr><th style='width:5%'>No</th><th style='width:25%'>Langkah</th><th>Penjelasan sederhana</th></tr>
    <tr><td>1</td><td><b>Mengumpulkan data</b></td><td>Jumlah penduduk, sekolah, guru, murid, puskesmas, dan apotek untuk 20 kecamatan tahun {TH0}–{TH1}.</td></tr>
    <tr><td>2</td><td><b>Menghitung indikator</b></td><td>Fasilitas dibagi jumlah penggunanya, misalnya sekolah per 10.000 penduduk dan murid per guru, agar kecamatan besar dan kecil bisa dibandingkan adil.</td></tr>
    <tr><td>3</td><td><b>Merata-rata 4 tahun</b></td><td>Supaya hasil tidak terpengaruh perubahan sesaat pada satu tahun.</td></tr>
    <tr><td>4</td><td><b>Menyamakan skala</b></td><td>Setiap indikator diubah ke skala yang sama (standardisasi), supaya tidak ada indikator yang mendominasi hanya karena angkanya besar.</td></tr>
    <tr><td>5</td><td><b>Mengelompokkan (K-Means)</b></td><td>Komputer mencari 4 kelompok kecamatan yang kondisi ketujuh indikatornya paling mirip satu sama lain.</td></tr>
    <tr><td>6</td><td><b>Menguji kestabilan</b></td><td>Pengelompokan diulang dengan data berbeda. Kecamatan yang hampir selalu masuk segmen yang sama disebut anggota inti.</td></tr>
    <tr><td>7</td><td><b>Memberi nama dan rekomendasi</b></td><td>Setiap segmen dinamai dari ciri paling menonjolnya, lalu disusun rekomendasi kebijakannya.</td></tr>
    </table></div>""", unsafe_allow_html=True)

    subjudul("Catatan penting")
    penjelasan(
        "<ul style='margin:0;padding-left:20px'>"
        "<li><b>Sumber:</b> Kecamatan Dalam Angka (KCDA) BPS Kabupaten Tuban, yang mengolah data Podes, Dapodik/EMIS, dan Dispendukcapil.</li>"
        "<li><b>Puskesmas</b> dihitung per 10.000 penduduk, bukan per desa, karena satu puskesmas memang melayani satu wilayah kecamatan.</li>"
        "<li><b>Data puskesmas dan apotek tahun 2022</b> disamakan dengan 2023, karena data 2022 dicatat dengan cara berbeda.</li>"
        "<li>Angka “± … penduduk” dibulatkan ke ratusan terdekat agar mudah dibaca.</li>"
        "<li><b>Segmen bukan peringkat.</b> Segmen menunjukkan jenis masalah yang mirip, bukan urutan baik-buruk.</li></ul>",
        "Yang perlu diketahui")

    subjudul("Unduh data hasil")
    unduh = rata[["Segmen", "Stabil_%", "Penduduk"] + FITUR].copy()
    for f in FITUR:
        unduh[IND[f]["ukuran"]] = unduh[f].map(lambda x, f=f: bulat(f, ubah(f, x)))
    unduh = unduh.drop(columns=FITUR).reset_index().rename(columns={"Stabil_%": "Kestabilan (%)", f"Penduduk": f"Penduduk {TH1}"})
    unduh["Kestabilan (%)"] = unduh["Kestabilan (%)"].round(0)
    st.download_button("⬇️ Unduh hasil segmentasi (CSV)", unduh.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                       file_name="hasil_segmentasi_kecamatan_tuban.csv", mime="text/csv")
    penjelasan("File CSV berisi segmen, kestabilan, jumlah penduduk, dan ketujuh indikator dalam ukuran mudah "
               "(penduduk per fasilitas, murid per guru, dan persen desa ber-apotek) untuk setiap kecamatan. "
               "File bisa dibuka di Excel.", "Isi file")

st.markdown(f"<div class='foot'>Projek 2 · Magang MBKM BPS Kabupaten Tuban · Program Studi Sains Data, UPN “Veteran” Jawa Timur · "
            f"<a href='{LINK_PETA}' target='_blank' style='color:{NAVY};font-weight:700'>Peta QGIS</a></div>",
            unsafe_allow_html=True)
