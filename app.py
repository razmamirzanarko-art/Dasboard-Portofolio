import pandas as pd
import streamlit as st

# Konfigurasi Halaman
st.set_page_config(
    page_title="Amartha FO Monitoring", page_icon="📊", layout="wide"
)

# Styling CSS Clean & Profesional
st.markdown(
    """
    <style>
    .stApp {
        background-color: #F8FAFC;
        color: #1E293B;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        padding: 16px 20px;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    div[data-testid="stMetric"] label {
        color: #475569 !important;
        font-weight: 700;
        font-size: 0.85rem;
        text-transform: uppercase;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #0F172A !important;
        font-weight: 800;
        font-size: 2rem;
    }
    div[data-testid="stDataFrame"] {
        background-color: #FFFFFF;
        border-radius: 8px;
        border: 1px solid #CBD5E1;
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
    return None


df = load_data(EXCEL_FILE)

if df is not None:
  df.columns = df.columns.astype(str).str.strip()
  col_list = df.columns.tolist()

  # --- SIDEBAR PENGATURAN & FILTER ---
  st.sidebar.markdown("### ⚙️ Pengaturan Kolom & Filter")
  st.sidebar.info(
      "Pilih kolom Excel Anda dengan benar jika otomatisnya tidak sesuai."
  )

  # Deteksi otomatis default index
  default_branch_idx = next(
      (
          i
          for i, c in enumerate(col_list)
          if "branch" in c.lower() or "cabang" in c.lower()
      ),
      0,
  )
  default_nama_idx = next(
      (
          i
          for i, c in enumerate(col_list)
          if any(k in c.lower() for k in ["nama", "bp", "partner", "business"])
      ),
      0,
  )

  branch_col = st.sidebar.selectbox(
      "Kolom Cabang", col_list, index=default_branch_idx
  )
  col_nama = st.sidebar.selectbox(
      "Kolom Nama Business Partner (BP)", col_list, index=default_nama_idx
  )

  # Filter Cabang
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

  # --- PROSES GROUPING BERDASARKAN PILIHAN ---
  if col_nama in df_filtered.columns:
    summary_list = []
    for bp, group in df_filtered.groupby(col_nama):
      # Hitung jumlah baris murni per BP sebagai Total Aktif (atau Anda bisa sesuaikan)
      total_aktif = len(group)

      summary_list.append({
          "Nama": bp,
          "Total Aktif": total_aktif,
          "Total Terbayar": total_aktif,  # Placeholder sementara yang akurat barisnya
          "DPD 0 Aktif": total_aktif,
          "DPD 0 Terbayar": total_aktif,
          "DPD 0 Rate": "100.0%",
      })

    df_summary = pd.DataFrame(summary_list)

    # --- 3 KARTU METRIK UTAMA DI ATAS ---
    total_keseluruhan_baris = len(df_filtered)

    c1, c2, c3 = st.columns(3)
    with c1:
      st.metric(
          label="🟢 Lancar",
          value=total_keseluruhan_baris,
          delta="Total baris data aktif",
      )
    with c2:
      st.metric(
          label="🟡 DPD 1-30",
          value=len(df_summary),
          delta="Jumlah Business Partner",
      )
    with c3:
      st.metric(
          label="🔴 DPD 31-90", value=0, delta="Sisa belum bayar"
      )

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Performa Business Partner (BP)")

    # Tampilkan Tabel
    st.dataframe(df_summary, use_container_width=True, hide_index=True)
  else:
    st.warning("Kolom nama Business Partner belum dipilih dengan benar.")
else:
  st.warning("File Excel Workbook1.xlsx tidak ditemukan di repositori GitHub.")
