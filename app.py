from datetime import datetime
import pandas as pd
import streamlit as st

# Konfigurasi Halaman
st.set_page_config(
    page_title="Amartha FO Monitoring", page_icon="📊", layout="wide"
)

# Styling CSS agar Mirip Persis Sistem Asli Amartha
st.markdown(
    """
    <style>
    .stApp {
        background-color: #F4F6F9;
        color: #333333;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    h1, h2, h3 {
        color: #2C3E50;
    }
    /* Kotak Metrik ala Amartha */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #D1D8DD;
        padding: 15px 20px;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    div[data-testid="stMetric"] label {
        color: #7F8C8D !important;
        font-weight: 600;
        font-size: 0.8rem;
        text-transform: uppercase;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #2C3E50 !important;
        font-weight: 700;
        font-size: 1.8rem;
    }
    /* Tabel Styling */
    div[data-testid="stDataFrame"] {
        background-color: #FFFFFF;
        border-radius: 6px;
        border: 1px solid #D1D8DD;
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
  if "Latest Payment Date" in df.columns:
    df["Latest Payment Date Parsed"] = pd.to_datetime(
        df["Latest Payment Date"], errors="coerce"
    )

  # Deteksi Kolom Cabang
  branch_col = None
  for col in df.columns:
    if "branch" in str(col).lower() or "cabang" in str(col).lower():
      branch_col = col
      break
  if not branch_col:
    branch_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]

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

  # --- HEADER UTAMA ---
  st.markdown(
      f"<p style='color: #7F8C8D; font-size: 0.85rem; margin-bottom:"
      " -10px;'>Home / Branches / FO Monitoring / Pembayaran</p>",
      unsafe_allow_html=True,
  )
  st.title(title_text)
  st.markdown(
      "<p style='color: #333333; font-weight: 600; font-size: 0.95rem;'>Minggu"
      " ini, 13 - 19 September 2026</p>",
      unsafe_allow_html=True,
  )
  st.markdown("<hr style='border: 1px solid #D1D8DD;'>", unsafe_allow_html=True)


  # --- KATEGORISASI STATUS ---
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

  # --- 3 KARTU UTAMA (Lancar, DPD 1-30, DPD 31-90) ---
  tot_lancar = len(df_filtered[df_filtered["Kat_Status"] == "DPD 0"])
  tot_dpd1 = len(df_filtered[df_filtered["Kat_Status"] == "DPD 1-30"])
  tot_dpd31 = len(df_filtered[df_filtered["Kat_Status"] == "DPD 31-90"])

  c1, c2, c3 = st.columns(3)
  with c1:
    st.metric(label="Lancar", value=tot_lancar, delta="Total pinjaman aktif")
  with c2:
    st.metric(label="DPD 1-30", value=tot_dpd1, delta="Total pinjaman aktif")
  with c3:
    st.metric(label="DPD 31-90", value=tot_dpd31, delta="Total pinjaman aktif")

  st.markdown("<br>", unsafe_allow_html=True)
  st.subheader("Performa Business Partner (BP)")

  # --- TABEL PERFORMA BP (Format Kolom Persis Sistem Amartha) ---
  bp_candidates = [
      col for col in df_filtered.columns if "bp" in col.lower() or "nama" in col.lower()
  ]
  bp_col = bp_candidates[0] if bp_candidates else df_filtered.columns[0]

  if bp_col in df_filtered.columns:
    bp_data = []
    for bp, group in df_filtered.groupby(bp_col):
      # Total Pinjaman
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

      bp_data.append({
          "Nama": bp,
          "Total Aktif": tot_aktif,
          "Total Terbayar": tot_terbayar,
          "DPD 0 Aktif": d0_aktif,
          "DPD 0 Terbayar": d0_terbayar,
          "DPD 0 Rate": f"{d0_rate}%",
          "DPD 1-30 Aktif": d1_aktif,
          "DPD 1-30 Terbayar": d1_terbayar,
          "DPD 1-30 Rate": f"{d1_rate}%",
          "DPD 31-90 Aktif": d31_aktif,
          "DPD 31-90 Terbayar": d31_terbayar,
          "DPD 31-90 Rate": f"{d31_rate}%",
      })

    df_table = pd.DataFrame(bp_data)
    st.dataframe(df_table, use_container_width=True, hide_index=True)
  else:
    st.warning("Kolom nama tidak ditemukan.")
else:
  st.warning("File Excel belum terbaca.")
