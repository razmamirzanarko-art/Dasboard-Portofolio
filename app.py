from datetime import datetime
import pandas as pd
import streamlit as st

# 1. Konfigurasi Halaman & Tema Elegan ala Perusahaan Asing
st.set_page_config(
    page_title="Amartha Executive Billing Dashboard",
    page_icon="💼",
    layout="wide",
)

# Styling CSS Tema Korporat Profesional (Clean, Modern, & Elegan)
st.markdown(
    """
    <style>
    .stApp {
        background-color: #F8FAFC;
        color: #1E293B;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    h1 {
        color: #0F172A;
        font-weight: 700;
        font-size: 2rem;
        letter-spacing: -0.025em;
    }
    h3 {
        color: #334155;
        font-weight: 600;
    }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        padding: 16px 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
    }
    div[data-testid="stMetric"] label {
        color: #64748B !important;
        font-weight: 500;
        font-size: 0.875rem;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #0F172A !important;
        font-weight: 700;
        font-size: 1.75rem;
    }
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }
    div[data-testid="stDataFrame"] {
        background-color: #FFFFFF;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
        padding: 8px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# 2. Load Data Excel (.xlsx) Terbaru Anda
EXCEL_FILE = (
    "Ops Report Penagihan 2026-09-19 (6).xlsx"  # Sesuaikan jika nama file beda
)


@st.cache_data
def load_data(file_path):
  try:
    if file_path.endswith(".xlsx"):
      return pd.read_excel(file_path)
    else:
      return pd.read_csv(file_path)
  except Exception as e:
    st.error(f"Gagal memuat file Excel: {e}")
    return None


df = load_data(EXCEL_FILE)

if df is not None:
  # Parsing tanggal pembayaran jika ada
  if "Latest Payment Date" in df.columns:
    df["Latest Payment Date Parsed"] = pd.to_datetime(
        df["Latest Payment Date"], errors="coerce"
    )

  # --- MENCARI KOLOM NAMA CABANG / WILAYAH TEKS ---
  # Mencari kolom yang mengandung kata 'branch' atau 'cabang' yang bertipe teks
  branch_col = None
  for col in df.columns:
    col_lower = str(col).lower()
    if (
        ("branch" in col_lower or "cabang" in col_lower)
        and "id" not in col_lower
        and df[col].dtype == "object"
    ):
      branch_col = col
      break

  # Jika tidak ketemu kolom teks khusus nama cabang, cari yang ada kata branch/cabang secara umum
  if not branch_col:
    for col in df.columns:
      if "branch" in str(col).lower() or "cabang" in str(col).lower():
        branch_col = col
        break

  # Fallback terakhir jika tetap tidak ketemu
  if not branch_col:
    branch_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]

  # --- SIDEBAR FILTER WILAYAH ---
  st.sidebar.markdown(
      "<h3 style='color: #0F172A; font-size: 1.1rem;'>⚙️ Filter Kontrol</h3>",
      unsafe_allow_html=True,
  )
  branch_list = ["Semua Cabang"] + sorted(
      df[branch_col].dropna().astype(str).unique().tolist()
  )
  selected_branch = st.sidebar.selectbox("Pilih Nama Cabang", branch_list)

  if selected_branch != "Semua Cabang":
    df_filtered = df[df[branch_col].astype(str) == selected_branch].copy()
    title_text = f"Executive Performance: Cabang {selected_branch}"
  else:
    df_filtered = df.copy()
    title_text = "Executive Performance: Keseluruhan Cabang"

  # --- HEADER UTAMA ---
  st.title(title_text)
  st.markdown(
      "<p style='color: #64748B; font-size: 1rem; margin-top: -10px;'>Periode"
      " Laporan: Minggu ini, 13 - 19 September 2026</p>",
      unsafe_allow_html=True,
  )
  st.markdown("<hr style='border: 1px solid #E2E8F0;'>", unsafe_allow_html=True)


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

  # --- 4 KARTU METRIK UTAMA ---
  tot_lancar_card = len(df_filtered[df_filtered["Kat_Status"] == "DPD 0"])
  tot_dpd1_card = len(df_filtered[df_filtered["Kat_Status"] == "DPD 1-30"])
  tot_dpd31_card = len(df_filtered[df_filtered["Kat_Status"] == "DPD 31-90"])
  tot_dpd90_card = len(df_filtered[df_filtered["Kat_Status"] == "DPD 90+"])

  col1, col2, col3, col4 = st.columns(4)
  with col1:
    st.metric(label="Lancar (Aktif)", value=tot_lancar_card)
  with col2:
    st.metric(label="DPD 1-30 (Aktif)", value=tot_dpd1_card)
  with col3:
    st.metric(label="DPD 31-90 (Aktif)", value=tot_dpd31_card)
  with col4:
    st.metric(label="DPD 90+ (Aktif)", value=tot_dpd90_card)

  st.markdown("<br>", unsafe_allow_html=True)
  st.subheader("Business Partner (BP) Performance Breakdown")

  # --- TABEL PERFORMA BUSINESS PARTNER (BP) ---
  bp_candidates = [
      col for col in df_filtered.columns if "bp" in col.lower() or "nama" in col.lower()
  ]
  bp_col = bp_candidates[0] if bp_candidates else df_filtered.columns[0]

  if bp_col in df_filtered.columns:
    bp_data = []
    for bp, group in df_filtered.groupby(bp_col):
      tot_aktif = len(group)
      tot_terbayar = int(group["Sudah_Bayar"].sum())

      # DPD 0
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
          "Business Partner": bp,
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
    st.warning("Kolom Business Partner tidak ditemukan dalam data.")
else:
  st.warning(
      "Silakan pastikan file Excel .xlsx terbaru Anda sudah di-upload dan"
      " namanya sesuai dengan variabel EXCEL_FILE."
  )
