# app.py (rev)
# -------------------------------------------------------------
# Streamlit App: Transform peserta Excel -> output per JENIS TES
# Perubahan utama dari versi sebelumnya:
# - Picker sheet diletakkan DI ATAS pratinjau masing-masing file (bukan di sidebar)
# - Pratinjau terpisah per JENIS_TES dalam bentuk TAB + badge jumlah baris
# - Label tombol download memuat jumlah data per jenis tes
# - Tambah tombol "Download All (ZIP)" sekali klik
# - Normalisasi tanggal lebih tangguh (string beragam + serial Excel)
# - Deteksi nama kolom fleksibel + opsi override manual via selectbox jika tidak ketemu
# - Perapihan pembuatan nama file & sanitasi karakter
# -------------------------------------------------------------

import io
import re
import zipfile
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generator Output Peserta", layout="wide")
st.title("📄 Generator Output Peserta per JENIS TES")

# -------------------------------------------------------------
# Sidebar: Nilai Default
# -------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Nilai Default Output")
    default_id_pendidikan = st.number_input("ID_PENDIDIKAN", value=45, step=1)
    default_id_jabatan = st.number_input("ID_JABATAN", value=10932, step=1)
    default_id_jenis_jabatan = st.number_input("ID_JENIS_JABATAN", value=4, step=1)

# -------------------------------------------------------------
# Helper
# -------------------------------------------------------------

def excel_sheet_picker(label: str):
    """Upload Excel + pilih sheet tepat di atas pratinjau."""
    up = st.file_uploader(label, type=["xlsx", "xls"])
    if up is None:
        return None, None, None
    xls = pd.ExcelFile(up)
    sheet = st.selectbox("Pilih sheet", xls.sheet_names, key=f"sheet_{label}")
    return up, xls, sheet

@st.cache_data(show_spinner=False)
def read_sheet(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, dtype=str)

@st.cache_data(show_spinner=False)
def build_ref_options(df_ref: pd.DataFrame) -> pd.DataFrame:
    # Expect columns: ID, NAMA (case-insensitive)
    cols = {c.lower(): c for c in df_ref.columns}
    id_col = cols.get("id")
    nama_col = cols.get("nama")
    if id_col is None or nama_col is None:
        st.error("File referensi harus memiliki kolom 'ID' dan 'NAMA'.")
        st.stop()
    ref = df_ref[[id_col, nama_col]].rename(columns={id_col: "ID", nama_col: "NAMA"}).copy()
    ref["ID"] = ref["ID"].astype(str).str.strip()
    ref["NAMA"] = ref["NAMA"].astype(str).str.strip()
    ref["OPTION"] = ref["ID"] + " — " + ref["NAMA"]
    return ref

@st.cache_data(show_spinner=False)
def normalize_date_scalar(x) -> str:
    """Normalisasi satu nilai tanggal ke yyyy-mm-dd. Tidak valid -> "" (kosong)."""
    if x is None:
        return ""
    # Sudah string kosong
    if isinstance(x, str) and x.strip() == "":
        return ""
    # Coba: to_datetime umum
    try:
        dt = pd.to_datetime(x, errors="raise", dayfirst=True)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    # Coba: angka serial Excel (umum: origin=1899-12-30)
    try:
        # Terima float/str yang tampak numerik
        if isinstance(x, (int, float)) or (isinstance(x, str) and re.fullmatch(r"\d+(\.\d+)?", x)):
            val = float(x)
            if not np.isnan(val):
                dt = pd.to_datetime(val, unit="D", origin="1899-12-30", errors="raise")
                return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    # Coba: format manual umum dd/mm/yyyy atau dd-mm-yyyy
    if isinstance(x, str):
        s = x.strip()
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%Y-%m-%d", "%d %b %Y", "%d %B %Y"):
            try:
                dt = pd.to_datetime(s, format=fmt)
                return dt.strftime("%Y-%m-%d")
            except Exception:
                continue
    return ""

@st.cache_data(show_spinner=False)
def normalize_date_series(s: pd.Series) -> pd.Series:
    return s.apply(normalize_date_scalar)

COLMAP_CANDIDATES = {
    "NO_PESERTA": ["NO_PESERTA", "NO PESERTA", "NIP", "NO"],
    "NAMA": ["NAMA", "NAMA PESERTA"],
    "TMPT_LAHIR": ["TMP_LAHIR", "TMPT_LAHIR", "TEMPAT LAHIR", "TMP LAHIR", "TEMPAT_LAHIR"],
    "TGL_LAHIR": ["TGL_LAHIR", "TANGGAL LAHIR", "TGL LAHIR", "TANGGAL_LAHIR"],
    "JENIS_TES": ["JENIS TES", "JENIS_TES"],
}


def pick_first_existing(df: pd.DataFrame, cands: List[str]) -> Optional[str]:
    for c in cands:
        if c in df.columns:
            return c
    return None


def sanitize_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]", "-", str(name))
    name = re.sub(r"\s+", " ", name).strip()
    return name or "output"

# -------------------------------------------------------------
# 1) Upload & pilih sheet PESERTA
# -------------------------------------------------------------
st.header("1) Upload File Peserta")
peserta_up, peserta_xls, peserta_sheet = excel_sheet_picker("Upload file peserta (Excel)")

if peserta_up is not None and peserta_sheet is not None:
    df_peserta = read_sheet(peserta_up.getvalue(), peserta_sheet)
    df_peserta.columns = [str(c).strip() for c in df_peserta.columns]
    st.caption(f"📑 Pratinjau Peserta — sheet: **{peserta_sheet}** (5 baris)")
    st.dataframe(df_peserta.head(5), use_container_width=True)
else:
    st.info("Unggah file peserta dan pilih sheet untuk melanjutkan.")

# -------------------------------------------------------------
# 2) Upload & pilih sheet REFERENSI INSTANSI
# -------------------------------------------------------------
st.header("2) Upload File Referensi Instansi")
ref_up, ref_xls, ref_sheet = excel_sheet_picker("Upload file referensi instansi (Excel)")

if ref_up is not None and ref_sheet is not None:
    df_ref_raw = read_sheet(ref_up.getvalue(), ref_sheet)
    df_ref_raw.columns = [str(c).strip() for c in df_ref_raw.columns]
    st.caption(f"📑 Pratinjau Referensi — sheet: **{ref_sheet}** (5 baris)")
    st.dataframe(df_ref_raw.head(5), use_container_width=True)
else:
    st.info("Unggah file referensi dan pilih sheet untuk melanjutkan.")

# -------------------------------------------------------------
# 3) Jika kedua file siap, lanjutkan mapping dan pembuatan output
# -------------------------------------------------------------
if (peserta_up is not None and peserta_sheet is not None) and (ref_up is not None and ref_sheet is not None):
    df_ref = build_ref_options(df_ref_raw)

    st.header("3) Pilih Instansi")
    picked_option = st.selectbox(
        "Cari & pilih instansi (ID — NAMA)",
        options=df_ref["OPTION"].tolist(),
        index=None,
        placeholder="Ketik untuk mencari...",
    )
    if picked_option is None:
        st.warning("Pilih instansi terlebih dahulu.")
        st.stop()

    picked_row = df_ref.loc[df_ref["OPTION"] == picked_option].iloc[0]
    picked_id = str(picked_row["ID"])  # ID_INSTANSI
    picked_nama = str(picked_row["NAMA"])  # UNIT_KERJA & INDUK

    # --- Kolom mapping otomatis + override manual --->
    st.header("4) Mapping Kolom Input → Standar Output")
    auto_map = {}
    for k, cands in COLMAP_CANDIDATES.items():
        auto_map[k] = pick_first_existing(df_peserta, cands)

    col_list = ["(pilih)"] + list(df_peserta.columns)
    c1, c2, c3, c4, c5 = st.columns(5)
    src_no = c1.selectbox("NO_PESERTA → NIP", col_list, index=col_list.index(auto_map["NO_PESERTA"]) if auto_map["NO_PESERTA"] else 0)
    src_nama = c2.selectbox("NAMA → NAMA", col_list, index=col_list.index(auto_map["NAMA"]) if auto_map["NAMA"] else 0)
    src_tmpl = c3.selectbox("TMP/TMPT_LAHIR → TMPT_LAHIR", col_list, index=col_list.index(auto_map["TMPT_LAHIR"]) if auto_map["TMPT_LAHIR"] else 0)
    src_tgll = c4.selectbox("TGL LAHIR → TGL_LAHIR", col_list, index=col_list.index(auto_map["TGL_LAHIR"]) if auto_map["TGL_LAHIR"] else 0)
    src_jenis = c5.selectbox("JENIS TES → JENIS_TES", col_list, index=col_list.index(auto_map["JENIS_TES"]) if auto_map["JENIS_TES"] else 0)

    required_map = {
        "NO_PESERTA": src_no if src_no != "(pilih)" else None,
        "NAMA": src_nama if src_nama != "(pilih)" else None,
        "TMPT_LAHIR": src_tmpl if src_tmpl != "(pilih)" else None,
        "TGL_LAHIR": src_tgll if src_tgll != "(pilih)" else None,
        "JENIS_TES": src_jenis if src_jenis != "(pilih)" else None,
    }

    missing = [k for k, v in required_map.items() if v is None]
    if missing:
        st.error("Kolom wajib belum lengkap: " + ", ".join(missing))
        st.stop()

    # --- Bangun DF kerja standar --->
    nip_series = df_peserta[required_map["NO_PESERTA"]].astype(str).str.strip()
    # Jaga leading zero NIP
    nip_series = nip_series.apply(lambda s: re.sub(r"\.0$", "", s))

    work = pd.DataFrame({
        "NIP": nip_series,
        "NAMA": df_peserta[required_map["NAMA"]].astype(str).str.strip(),
        "TMPT_LAHIR": df_peserta[required_map["TMPT_LAHIR"]].astype(str).str.strip(),
        "TGL_LAHIR": normalize_date_series(df_peserta[required_map["TGL_LAHIR"]]),
        "JENIS_TES": df_peserta[required_map["JENIS_TES"]].astype(str).str.strip(),
    })

    work["UNIT_KERJA"] = picked_nama
    work["UNIT_KERJA_INDUK"] = picked_nama
    work["ID_INSTANSI"] = picked_id
    work["ID_PENDIDIKAN"] = str(int(default_id_pendidikan))
    work["ID_JABATAN"] = str(int(default_id_jabatan))
    work["ID_JENIS_JABATAN"] = str(int(default_id_jenis_jabatan))

    # Reorder
    out_cols = [
        "NIP", "NAMA", "TMPT_LAHIR", "TGL_LAHIR", "UNIT_KERJA", "UNIT_KERJA_INDUK",
        "ID_INSTANSI", "ID_PENDIDIKAN", "ID_JABATAN", "ID_JENIS_JABATAN", "JENIS_TES",
    ]
    work = work[out_cols]

    st.header("5) Pratinjau Hasil Per Jenis Tes + Unduhan")
    if work["JENIS_TES"].isna().all() or (work["JENIS_TES"].str.len() == 0).all():
        st.info("Kolom JENIS_TES kosong — tidak ada pembagian.")
        st.stop()

    groups: Dict[str, pd.DataFrame] = {
        k: v.drop(columns=["JENIS_TES"]).reset_index(drop=True)
        for k, v in work.groupby("JENIS_TES", dropna=False)
    }

    # Ringkasan jumlah per jenis tes
    counts_df = pd.DataFrame({
        "JENIS_TES": list(groups.keys()),
        "JUMLAH": [len(df) for df in groups.values()],
    })
    st.markdown("#### Ringkasan Jumlah per Jenis Tes")
    st.dataframe(counts_df, use_container_width=True)

    # Tabs per jenis tes
    tab_labels = [f"{str(k)} ({len(df)})" for k, df in groups.items()]
    tabs = st.tabs(tab_labels)

    # Untuk ZIP all
    zip_buffer = io.BytesIO()
    zip_file = zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED)

    for (jenis, df_out), tab in zip(groups.items(), tabs):
        with tab:
            st.markdown(f"### Jenis Tes: **{jenis}** · Jumlah: **{len(df_out)}**")
            st.dataframe(df_out.head(10), use_container_width=True)

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                df_out.to_excel(writer, index=False)
            buf.seek(0)

            safe_inst = sanitize_filename(picked_nama)
            safe_jenis = sanitize_filename(jenis)
            filename = f"{safe_inst}_{safe_jenis}.xlsx"

            st.download_button(
                label=f"⬇️ Download ({len(df_out)} baris)",
                data=buf.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{safe_jenis}",
            )

            # Tambahkan ke ZIP
            zip_file.writestr(filename, buf.getvalue())

    zip_file.close()
    zip_buffer.seek(0)

    st.divider()
    st.subheader("📦 Download All (ZIP)")
    st.caption("Satu klik untuk mengunduh seluruh file hasil per JENIS_TES.")
    st.download_button(
        label=f"⬇️ Download All ZIP ({len(groups)} file)",
        data=zip_buffer.getvalue(),
        file_name=f"{sanitize_filename(picked_nama)}_ALL_JENIS_TES.zip",
        mime="application/zip",
        key="dl_all_zip",
    )
else:
    st.stop()
