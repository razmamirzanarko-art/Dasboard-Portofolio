import pandas as pd
import streamlit as st

# Pengaturan halaman Streamlit
st.set_page_config(
    page_title="Dashboard Performa Penagihan", page_icon="📊", layout="wide"
)

# Nama file data di repository GitHub (sesuai format CSV)
EXCEL_FILE = "Ops Report Penagihan 2026-09-19 (6).csv"


@st.cache_data
def load_data(file_path):
  try:
    # Membaca file CSV
    df = pd.read_csv(file_path)
    return df
  except Exception as e:
    return None


# Memuat data
df = load_data(EXCEL_FILE)

if df is not None:
  # Tampilan Header Utama
  st.title("📊 Dashboard Performa Penagihan")
  st.markdown("*Periode: Minggu ini, 13 - 19 September 2026*")
  st.markdown("---")

  # Sidebar untuk Filter Cabang / Area
  st.sidebar.header("🔍 Filter Wilayah")

  # Deteksi kolom cabang/area secara otomatis dari data CSV Anda
  # (Biasanya kolom berisi nama cabang/area)
  possible_cols = [
      col
      for col in df.columns
      if "cabang" in col.lower()
      or "area" in col.lower()
      or "branch" in col.lower()
  ]

  if possible_cols:
    selected_col = possible_cols[0]
    branches = ["Semua Cabang / Area"] + list(df[selected_col].dropna().unique())
    selected_branch = st.sidebar.selectbox("Pilih Cabang / Area", branches)

    if selected_branch != "Semua Cabang / Area":
      filtered_df = df[df[selected_col] == selected_branch]
      st.subheader(f"Performa: {selected_branch}")
    else:
      filtered_df = df
      st.subheader("Performa Keseluruhan (Semua Cabang & Area)")
  else:
    filtered_df = df
    selected_branch = "Semua"

  # Menampilkan Ringkasan Metrik (KPI) jika ada kolom angka yang sesuai
  # (Contoh: Total Tagihan, Tertagih, dll.)
  st.markdown("### 📈 Ringkasan Data")
  st.dataframe(filtered_df, use_container_width=True)

  # Tambahan informasi unduh data jika diperlukan
  st.download_button(
      label="📥 Unduh Data Terfilter (CSV)",
      data=filtered_df.to_csv(index=False).encode("utf-8"),
      file_name="laporan_penagihan_terfilter.csv",
      mime="text/csv",
  )

else:
  st.error(f"Gagal memuat file: {EXCEL_FILE}")
  st.info(
      "Silakan pastikan nama file CSV di repository GitHub Anda sudah sama"
      f" persis dengan {EXCEL_FILE}."
  )
