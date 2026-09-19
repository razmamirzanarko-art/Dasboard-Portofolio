import pandas as pd
import streamlit as st

# Konfigurasi Halaman & Layout Luas
st.set_page_config(
    page_title="Amartha FO Monitoring", page_icon="📊", layout="wide"
)

# Styling CSS Agar Mirip Persis Sistem Asli Amartha
st.markdown(
    """
    <style>
    .stApp {
        background-color: #F8FAFC;
        color: #1E293B;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        padding: 16px 20px;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    div[data-testid="stMetric"] label {
        color: #64748B !important;
        font-weight: 600;
        font-size: 0.8rem;
        text-transform: uppercase;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #0F172A !important;
        font-weight: 700;
        font-size: 1.75rem;
    }
    div[data-testid="stDataFrame"] {
        background-color: #FFFFFF;
        border-radius: 8px;
        border: 1px solid #E2E8F0;
        padding: 4px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Load Data Excel
EXCEL_FILE = "Workbook1.xlsx"


@st.cache_data
def load_data(file_path):
  try:
    return pd.read_excel(file_path)
  except Exception as e:
    st.error(f"Gagal memuat file: {e}")
    return None


df = load_data(EXCEL_FILE)

if df is not None:
  # Deteksi Kolom Cabang
  branch_col = next(
      (
          c
          for c in df.columns
          if "branch" in str(c).lower() or "cabang" in str(c).lower()
      ),
      df.columns[3] if len(df.columns) > 3 else df.columns[0],
  )

  # --- SIDEBAR FILTER ---
  st.sidebar.markdown("### ⚙️ Filter Cabang")
  branch_list = ["Semua Cabang"] + sorted(
      df[branch_col].dropna().astype(str).unique().tolist()
  )
  selected_branch = st.sidebar.selectbox("Pilih Cabang", branch_list)

  if selected_branch != "Semua Cabang":
    df_filtered = df[df[branch_col].astype(str) == selected_branch].copy()
    title_text = f"Performa: {selected_branch}"
  else:
    df_filtered = df.copy()
    title_text = "Performa: Keseluruhan Cabang"

  # --- HEADER UTAMA (PERSIS SEPERTI GAMBAR) ---
  st.markdown(
      "<p style='color: #64748B; font-size: 0.85rem; margin-bottom:"
      " -10px;'>Home / Branches / FO Monitoring / Pembayaran</p>",
      unsafe_allow_html=True,
  )
  st.title(title_text)
  st.markdown(
      "<p style='color: #334155; font-weight: 600; font-size: 0.95rem;'>Minggu"
      " ini, 13 - 19 September 2026</p>",
      unsafe_allow_html=True,
  )
  st.markdown("<hr style='border: 1px solid #E2E8F0;'>", unsafe_allow_html=True)

  # --- 3 KARTU UTAMA DI ATAS ---
  tot_lancar = len(
      df_filtered
  )  # Ganti dengan perhitungan metrik utama sesuai data Anda
  tot_dpd1 = int(len(df_filtered) * 0.1)
  tot_dpd31 = int(len(df_filtered) * 0.05)

  c1, c2, c3 = st.columns(3)
  with c1:
    st.metric(label="Lancar", value=tot_lancar, delta="Total pinjaman aktif")
  with c2:
    st.metric(label="DPD 1-30", value=tot_dpd1, delta="Total pinjaman aktif")
  with c3:
    st.metric(label="DPD 31-90", value=tot_dpd31, delta="Total pinjaman aktif")

  st.markdown("<br>", unsafe_allow_html=True)
  st.subheader("Performa Business Partner (BP)")

  # --- TABEL UTAMA (MENAMPILKAN APA ADANYA DARI EXCEL AGAR TIDAK SALAH) ---
  bp_candidates = [
      col for col in df_filtered.columns if "bp" in col.lower() or "nama" in col.lower()
  ]
  bp_col = bp_candidates[0] if bp_candidates else df_filtered.columns[0]

  if bp_col in df_filtered.columns:
    # Langsung tampilkan dataframe yang dibersihkan kolomnya agar persis tabel aslinya
    st.dataframe(df_filtered, use_container_width=True, hide_index=True)
  else:
    st.warning("Kolom nama tidak ditemukan.")
else:
  st.warning("File Excel belum terbaca.")
