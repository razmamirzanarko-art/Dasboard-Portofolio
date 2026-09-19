from datetime import datetime
import pandas as pd
import streamlit as st

# 1. Konfigurasi Halaman Dashboard
st.set_page_config(
    page_title="Dashboard Performa Penagihan - Amartha",
    page_icon="📊",
    layout="wide",
)

# 2. Load Data (Sesuaikan nama file excel Anda jika berbeda)
EXCEL_FILE = "Ops Report Penagihan 2026-09-19 (6).csv"  # Ganti dengan nama file terbaru Anda jika formatnya .xlsx atau .csv


@st.cache_data
def load_data(file_path):
  try:
    import csv
    import io
    import openpyxl

    if file_path.endswith(".xlsx"):
      wb = openpyxl.load_workbook(file_path)
      sheet = wb.active
      all_data = []
      for row in sheet.iter_rows(values_only=True):
        if row[0] is not None:
          r_parsed = next(csv.reader(io.StringIO(str(row[0]))))
          all_data.append(r_parsed)
      return pd.DataFrame(all_data[1:], columns=all_data[0])
    else:
      return pd.read_csv(file_path)
  except Exception as e:
    st.error(f"Gagal memuat file: {e}")
    return None


df = load_data(EXCEL_FILE)

if df is not None:
  # Parsing tanggal pembayaran jika ada
  if "Latest Payment Date" in df.columns:
    df["Latest Payment Date Parsed"] = pd.to_datetime(
        df["Latest Payment Date"], errors="coerce"
    )

  # --- SIDEBAR FILTER CABANG ---
  st.sidebar.header("🔍 Filter Wilayah")
  branch_col = "Branch" if "Branch" in df.columns else df.columns[3]

  branch_list = ["Semua Cabang"] + sorted(
      df[branch_col].dropna().astype(str).unique().tolist()
  )
  selected_branch = st.sidebar.selectbox("Pilih Cabang", branch_list)

  if selected_branch != "Semua Cabang":
    df_filtered = df[df[branch_col].astype(str) == selected_branch].copy()
    title_text = f"Performa: Cabang {selected_branch}"
  else:
    df_filtered = df.copy()
    title_text = "Performa: Keseluruhan Cabang"

  # --- HEADER UTAMA ---
  st.title(title_text)
  st.markdown("### Minggu ini, 13 - 19 September 2026")
  st.markdown("---")


  # --- FUNGSI KATEGORISASI STATUS & PEMBAYARAN ---
  def classify_row(row):
    s = str(row["Payment Status"]).upper()
    dt = (
        row["Latest Payment Date Parsed"]
        if "Latest Payment Date Parsed" in row and not pd.isna(row["Latest Payment Date Parsed"])
        else None
    )
    is_bayar = dt is not None and dt.month == 9 and (14 <= dt.day <= 19)

    if "LANCAR" in s or "DPD 0" in s:
      return "DPD 0", is_bayar
    elif any(x in s for x in ["DPD 1-7", "DPD 8-14", "DPD 15-30", "DPD 1-30"]):
      return "DPD 1-30", is_bayar
    elif any(
        x in s
        for x in [
            "DPD 31",
            "DPD 38",
            "DPD 54",
            "DPD 61",
            "DPD 68",
            "DPD 84",
            "DPD 31-90",
        ]
    ):
      return "DPD 31-90", is_bayar
    elif any(x in s for x in ["DPD 90", "DPD 90+"]):
      return "DPD 90+", is_bayar
    return "Lainnya", is_bayar


  res = df_filtered.apply(classify_row, axis=1)
  df_filtered["Kat_Status"] = [r[0] for r in res]
  df_filtered["Sudah_Bayar"] = [r[1] for r in res]

  # --- KARTU METRIK ATAS (4 KOTAK UTAMA SESUAI GAMBAR) ---
  tot_lancar_card = len(df_filtered[df_filtered["Kat_Status"] == "DPD 0"])
  tot_dpd1_card = len(df_filtered[df_filtered["Kat_Status"] == "DPD 1-30"])
  tot_dpd31_card = len(df_filtered[df_filtered["Kat_Status"] == "DPD 31-90"])
  tot_dpd90_card = len(df_filtered[df_filtered["Kat_Status"] == "DPD 90+"])

  col1, col2, col3, col4 = st.columns(4)
  with col1:
    st.metric(label="Lancar (Total Pinjaman Aktif)", value=tot_lancar_card)
  with col2:
    st.metric(label="DPD 1-30 (Total Pinjaman Aktif)", value=tot_dpd1_card)
  with col3:
    st.metric(label="DPD 31-90 (Total Pinjaman Aktif)", value=tot_dpd31_card)
  with col4:
    st.metric(label="DPD 90+ (Total Pinjaman Aktif)", value=tot_dpd90_card)

  st.markdown("---")
  st.subheader("Performa Business Partner (BP)")

  # --- TABEL PERFORMA BUSINESS PARTNER (BP) LENGKAP ---
  if "BP" in df_filtered.columns:
    bp_data = []
    for bp, group in df_filtered.groupby("BP"):
      # Total Pinjaman
      tot_aktif = len(group)
      tot_terbayar = int(group["Sudah_Bayar"].sum())

      # DPD 0 (Lancar)
      d0_grp = group[group["Kat_Status"] == "DPD 0"]
      d0_aktif = len(d0_grp)
      d0_terbayar = int(d0_grp["Sudah_Bayar"].sum())
      d0_rate = round((d0_terbayar / d0_aktif * 100), 1) if d0_aktif > 0 else 0.0

      # DPD 1-30
      d1_grp = group[group["Kat_Status"] == "DPD 1-30"]
      d1_aktif = len(d1_grp)
      d1_terbayar = int(d1_grp["Sudah_Bayar"].sum())
      d1_rate = round((d1_terbayar / d1_aktif * 100), 1) if d1_aktif > 0 else 0.0

      # DPD 31-90
      d31_grp = group[group["Kat_Status"] == "DPD 31-90"]
      d31_aktif = len(d31_grp)
      d31_terbayar = int(d31_grp["Sudah_Bayar"].sum())
      d31_rate = (
          round((d31_terbayar / d31_aktif * 100), 1) if d31_aktif > 0 else 0.0
      )

      # DPD 90+
      d90_grp = group[group["Kat_Status"] == "DPD 90+"]
      d90_aktif = len(d90_grp)
      d90_terbayar = int(d90_grp["Sudah_Bayar"].sum())
      d90_rate = (
          round((d90_terbayar / d90_aktif * 100), 1) if d90_aktif > 0 else 0.0
      )

      bp_data.append({
          "Nama": bp,
          "Total Aktif": tot_aktif,
          "Total Terbayar": tot_terbayar,
          "DPD 0 Aktif": d0_aktif,
          "DPD 0 Terbayar": d0_terbayar,
          "DPD 0 Rate (%)": f"{d0_rate}%",
          "DPD 1-30 Aktif": d1_aktif,
          "DPD 1-30 Terbayar": d1_terbayar,
          "DPD 1-30 Rate (%)": f"{d1_rate}%",
          "DPD 31-90 Aktif": d31_aktif,
          "DPD 31-90 Terbayar": d31_terbayar,
          "DPD 31-90 Rate (%)": f"{d31_rate}%",
          "DPD 90+ Aktif": d90_aktif,
          "DPD 90+ Terbayar": d90_terbayar,
          "DPD 90+ Rate (%)": f"{d90_rate}%",
      })

    summary_table = pd.DataFrame(bp_data)
    st.dataframe(summary_table, use_container_width=True, hide_index=True)
  else:
    st.warning("Kolom 'BP' tidak ditemukan di dalam data.")
else:
  st.warning("Silakan periksa kembali nama file Excel / CSV Anda di variabel EXCEL_FILE.")
