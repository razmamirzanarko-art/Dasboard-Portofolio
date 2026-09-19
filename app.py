from datetime import datetime
import pandas as pd
import streamlit as st

# 1. Konfigurasi Halaman Dashboard
st.set_page_config(
    page_title="Dashboard Performa Penagihan", page_icon="📊", layout="wide"
)

st.title("Performa: Astambul")
st.markdown("### Minggu ini, 13 - 19 September 2026")

# 2. Fungsi Load Data Excel
EXCEL_FILE = "Ops Report Penagihan 2026-09-19 (6).xlsx"


@st.cache_data
def load_data(file_path):
  try:
    import csv
    import io
    import openpyxl

    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    all_data = []
    for row in sheet.iter_rows(values_only=True):
      if row[0] is not None:
        val = str(row[0])
        r_parsed = next(csv.reader(io.StringIO(val)))
        all_data.append(r_parsed)

    header = all_data[0]
    df = pd.DataFrame(all_data[1:], columns=header)
    return df
  except Exception as e:
    st.error(f"Gagal memuat file Excel: {e}")
    return None


df = load_data(EXCEL_FILE)

if df is not None:
  # Konversi Latest Payment Date ke format datetime untuk filter tgl 14-19
  df["Latest Payment Date Parsed"] = pd.to_datetime(
      df["Latest Payment Date"], errors="coerce"
  )

  # Filter pembayaran di rentang tanggal 14 - 19 September 2026
  # (Sesuai ketentuan: acuan bayar di tgl 14-19 untuk Lancar, DPD 1-30, DPD 31-90)
  mask_tgl = (df["Latest Payment Date Parsed"].dt.month == 9) & (
      df["Latest Payment Date Parsed"].dt.day.between(14, 19)
  )
  df_tgl_14_19 = df[mask_tgl]

  # --- 3. KATEGORISASI STATUS PEMBAYARAN ---
  # Lancar / DPD 0
  lancar_df = df_tgl_14_19[
      df_tgl_14_19["Payment Status"].str.contains(
          "Lancar|DPD 0", case=False, na=False
      )
  ]
  # DPD 1-30
  dpd_1_30_df = df_tgl_14_19[
      df_tgl_14_19["Payment Status"].str.contains("DPD 1-30", case=False, na=False)
  ]
  # DPD 31-90
  dpd_31_90_df = df_tgl_14_19[
      df_tgl_14_19["Payment Status"].str.contains(
          "DPD 31-90", case=False, na=False
      )
  ]

  # --- 4. TAMPILAN KARTU METRIK UTAMA (TOP CARDS) ---
  col1, col2, col3, col4 = st.columns(4)

  with col1:
    st.metric(
        label="Lancar (Total Pinjaman Aktif)",
        value=len(lancar_df),
        help="Pembayaran lunas/lancar dengan acuan tgl 14-19",
    )
  with col2:
    st.metric(
        label="DPD 1-30 (Total Pinjaman Aktif)",
        value=len(dpd_1_30_df),
        help="DPD 1-30 dengan acuan pembayaran tgl 14-19",
    )
  with col3:
    st.metric(
        label="DPD 31-90 (Total Pinjaman Aktif)",
        value=len(dpd_31_90_df),
        help="DPD 31-90 dengan acuan pembayaran tgl 14-19",
    )
  with col4:
    st.metric(
        label="Total Baris Data (Keseluruhan)",
        value=len(df),
        help="Total keseluruhan data pada laporan",
    )

  st.markdown("---")

  # --- 5. TABEL PERFORMA BUSINESS PARTNER (BP) ---
  st.subheader("Performa Business Partner (BP)")

  if "BP" in df.columns:
    # Agregasi data per Business Partner (BP)
    bp_summary = (
        df.groupby("BP")
        .agg(
            Total_Pinjaman=("Loan ID", "count"),
            Total_Outstanding=("Outstanding", lambda x: pd.to_numeric(x, errors="coerce").sum()),
        )
        .reset_index()
    )

    st.markdown(
        "Tabel ringkasan performa berdasarkan *Business Partner (BP)* dari file"
        " Excel:"
    )
  
    # Filter pencarian BP opsional
    search_bp = st.text_input("Cari Nama Business Partner (BP):")
    if search_bp:
      bp_summary = bp_summary[
          bp_summary["BP"].str.contains(search_bp, case=False, na=False)
      ]

    st.dataframe(bp_summary, use_container_width=True)
  else:
    st.warning("Kolom 'BP' tidak ditemukan di dalam file Excel.")

  # --- 6. TABEL DETAIL DATA ---
  with st.expander("Lihat Detail Data Mentah (Filtered Tanggal 14-19)"):
    st.dataframe(df_tgl_14_19, use_container_width=True)

else:
  st.info(
      "Silakan pastikan file Excel *'Ops Report Penagihan 2026-09-19 (6).xlsx'*"
      " sudah berada di dalam repository GitHub yang sama dengan file"
      " app.py."
  )
