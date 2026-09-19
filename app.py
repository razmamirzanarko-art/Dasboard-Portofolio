from datetime import datetime
import pandas as pd
import streamlit as st

# 1. Konfigurasi Halaman Dashboard
st.set_page_config(
    page_title="Dashboard Performa Penagihan - Amartha",
    page_icon="📊",
    layout="wide",
)

st.title("Performa: Astambul")
st.markdown("### Minggu ini, 13 - 19 September 2026")
st.markdown("---")

# 2. Fungsi Load Data
EXCEL_FILE = "Ops Report Penagihan 2026-09-19 (6).csv"


@st.cache_data
def load_data(file_path):
  try:
    import csv
    import io
    import openpyxl

    # Cek apakah file xlsx atau csv
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
  # Konversi tanggal pembayaran ke datetime untuk filter tgl 14-19
  df["Latest Payment Date Parsed"] = pd.to_datetime(
      df["Latest Payment Date"], errors="coerce"
  )

  # Filter acuan tanggal pembayaran di rentang 14 - 19 September 2026
  mask_tgl = (df["Latest Payment Date Parsed"].dt.month == 9) & (
      df["Latest Payment Date Parsed"].dt.day.between(14, 19)
  )
  df_tgl = df[mask_tgl].copy()


  # Mapping kategori status pembayaran
  def categorize_status(status):
    if pd.isna(status):
      return "Lainnya"
    s = str(status).strip().upper()
    if "LANCAR" in s or "DPD 0" in s:
      return "Lancar"
    elif any(
        x in s
        for x in [
            "DPD 1-7",
            "DPD 8-14",
            "DPD 15-30",
            "DPD 1-30",
            "DPD 31",
            "DPD 38",
        ]
    ):
      # Kita kelompokkan sesuai rentang DPD di tabel referensi
      pass
    return s


  # --- 3. KARTU METRIK UTAMA (TOP CARDS) ---
  total_lancar = len(
      df[df["Payment Status"].str.contains("Lancar|DPD 0", case=False, na=False)]
  )
  total_dpd_1_30 = len(
      df[
          df["Payment Status"].str.contains(
              "DPD 1-7|DPD 8-14|DPD 15-30|DPD 1-30", case=False, na=False
          )
      ]
  )
  total_dpd_31_90 = len(
      df[
          df["Payment Status"].str.contains(
              "DPD 31|DPD 38|DPD 54|DPD 61|DPD 68|DPD 84|DPD 31-90",
              case=False,
              na=False,
          )
      ]
  )

  col1, col2, col3 = st.columns(3)
  with col1:
    st.metric(label="Lancar (Total Pinjaman Aktif)", value=total_lancar)
  with col2:
    st.metric(label="DPD 1-30 (Total Pinjaman Aktif)", value=total_dpd_1_30)
  with col3:
    st.metric(label="DPD 31-90 (Total Pinjaman Aktif)", value=total_dpd_31_90)

  st.markdown("---")
  st.subheader("Performa Business Partner (BP)")

  # --- 4. TABEL PERFORMA BUSINESS PARTNER (BP) DENGAN PERSENTASE ---
  if "BP" in df.columns:
    # Buat kolom bantu untuk pengelompokan status
    def get_group_status(row):
      s = str(row["Payment Status"]).upper()
      # Cek apakah sudah bayar di tgl 14-19
      dt = row["Latest Payment Date Parsed"]
      is_bayar_14_19 = (
          not pd.isna(dt) and dt.month == 9 and (14 <= dt.day <= 19)
      )

      if "LANCAR" in s or "DPD 0" in s:
        return "Lancar", is_bayar_14_19
      elif any(
          x in s for x in ["DPD 1-7", "DPD 8-14", "DPD 15-30", "DPD 1-30"]
      ):
        return "DPD 1-30", is_bayar_14_19
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
        return "DPD 31-90", is_bayar_14_19
      return "Lainnya", False

    res = df.apply(get_group_status, axis=1)
    df["Kat_Status"] = [r[0] for r in res]
    df["Sudah_Bayar_14_19"] = [r[1] for r in res]

    # Agregasi per Business Partner (BP)
    bp_grouped = []
    for bp, group in df.groupby("BP"):
      # Total Pinjaman Aktif & Terbayar keseluruhan
      tot_aktif = len(group)
      tot_terbayar = group["Sudah_Bayar_14_19"].sum()

      # Kategori Lancar / DPD 0
      lancar_grp = group[group["Kat_Status"] == "Lancar"]
      lancar_aktif = len(lancar_grp)
      lancar_terbayar = lancar_grp["Sudah_Bayar_14_19"].sum()
      lancar_rate = (
          (lancar_terbayar / lancar_aktif * 100) if lancar_aktif > 0 else 0.0
      )

      # Kategori DPD 1-30
      dpd1_grp = group[group["Kat_Status"] == "DPD 1-30"]
      dpd1_aktif = len(dpd1_grp)
      dpd1_terbayar = dpd1_grp["Sudah_Bayar_14_19"].sum()
      dpd1_rate = (
          (dpd1_terbayar / dpd1_aktif * 100) if dpd1_aktif > 0 else 0.0
      )

      # Kategori DPD 31-90
      dpd31_grp = group[group["Kat_Status"] == "DPD 31-90"]
      dpd31_aktif = len(dpd31_grp)
      dpd31_terbayar = dpd31_grp["Sudah_Bayar_14_19"].sum()
      dpd31_rate = (
          (dpd31_terbayar / dpd31_aktif * 100) if dpd31_aktif > 0 else 0.0
      )

      bp_grouped.append({
          "Nama BP": bp,
          "Total Aktif": tot_aktif,
          "Total Terbayar": tot_terbayar,
          "Lancar Aktif": lancar_aktif,
          "Lancar Terbayar": lancar_terbayar,
          "Lancar Rate (%)": round(lancar_rate, 1),
          "DPD 1-30 Aktif": dpd1_aktif,
          "DPD 1-30 Terbayar": dpd1_terbayar,
          "DPD 1-30 Rate (%)": round(dpd1_rate, 1),
          "DPD 31-90 Aktif": dpd31_aktif,
          "DPD 31-90 Terbayar": dpd31_terbayar,
          "DPD 31-90 Rate (%)": round(dpd31_rate, 1),
      })

    summary_df = pd.DataFrame(bp_grouped)

    # Tampilkan tabel interaktif menyerupai dashboard Amartha
    st.dataframe(summary_df, use_container_width=True)
  else:
    st.warning("Kolom 'BP' tidak ditemukan di dalam data.")

else:
  st.error(
      f"File {EXCEL_FILE} tidak ditemukan di repository GitHub Anda."
      " Pastikan file berekstensi .csv atau .xlsx sudah di-upload."
  )
