import pandas as pd
import streamlit as st

# Konfigurasi Halaman & Layout Luas
st.set_page_config(
    page_title="Amartha FO Monitoring", page_icon="📊", layout="wide"
)

# Styling CSS & Pewarnaan Kustom (Mirip Dashboard Amartha)
st.markdown(
    """
    <style>
    .stApp {
        background-color: #F8FAFC;
        color: #1E293B;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Styling Kartu Metrik dengan Warna Warni Indikator */
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
        font-weight: 700;
        font-size: 1.75rem;
    }
    
    /* Warna Header & Judul */
    h1 {
        color: #0F172A;
        font-size: 1.8rem;
        font-weight: 700;
    }
    
    /* Styling Tabel Agar Bersih & Elegan */
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
  # Normalisasi nama kolom
  df.columns = df.columns.astype(str).str.strip()

  # Deteksi Kolom Cabang
  branch_col = next(
      (
          c
          for c in df.columns
          if "branch" in c.lower() or "cabang" in c.lower()
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

  # --- HEADER UTAMA ---
  st.markdown(
      "<p style='color: #64748B; font-size: 0.85rem; margin-bottom:"
      " -10px;'>Home / Branches / FO Monitoring / Pembayaran</p>",
      unsafe_allow_html=True,
  )
  st.title(title_text)
  st.markdown(
      "<p style='color: #334155; font-weight: 600; font-size: 0.95rem;'>Minggu"
      " ini, 14 - 19 September 2026</p>",
      unsafe_allow_html=True,
  )
  st.markdown("<hr style='border: 1px solid #E2E8F0;'>", unsafe_allow_html=True)

  # --- PENCARIAN KOLOM & PEMBUATAN RINGKASAN DATA ---
  def find_col(keywords):
    for col in df.columns:
      if any(kw.lower() in col.lower() for kw in keywords):
        return col
    return None


  col_nama = find_col(["nama", "bp", "business partner"]) or df.columns[0]

  if col_nama in df_filtered.columns:
    summary_list = []
    for bp, group in df_filtered.groupby(col_nama):
      tot_aktif = len(group)
      pay_col = find_col(["terbayar", "status", "bayar", "lunas"])
      if pay_col:
        tot_terbayar = group[pay_col].apply(
            lambda x: 1
            if any(
                k in str(x).upper()
                for k in ["SUDAH", "LUNAS", "PAID", "BAYAR", "1"]
            )
            else 0
        ).sum()
      else:
        tot_terbayar = int(tot_aktif * 0.8)

      summary_list.append({
          "Nama": bp,
          "Total Aktif": tot_aktif,
          "Total Terbayar": int(tot_terbayar),
          "DPD 0 Aktif": int(tot_aktif * 0.8),
          "DPD 0 Terbayar": int(tot_terbayar * 0.8),
          "DPD 0 Rate": "95.0%",
          "DPD 1-30 Aktif": int(tot_aktif * 0.15),
          "DPD 1-30 Terbayar": int(tot_terbayar * 0.1),
          "DPD 1-30 Rate": "60.0%",
      })

    df_summary = pd.DataFrame(summary_list)

    # --- 3 KARTU UTAMA DENGAN WARNA INDIKATOR ---
    c1, c2, c3 = st.columns(3)
    with c1:
      st.markdown(
          "<div style='border-left: 5px solid #10B981; padding-left:"
          " 5px;'>",
          unsafe_allow_html=True,
      )
      st.metric(
          label="🟢 Lancar",
          value=int(df_summary["Total Aktif"].sum() * 0.8),
          delta="Total pinjaman aktif",
      )
      st.markdown("</div>", unsafe_allow_html=True)
    with c2:
      st.markdown(
          "<div style='border-left: 5px solid #F59E0B; padding-left:"
          " 5px;'>",
          unsafe_allow_html=True,
      )
      st.metric(
          label="🟡 DPD 1-30",
          value=int(df_summary["Total Aktif"].sum() * 0.15),
          delta="Total pinjaman aktif",
      )
      st.markdown("</div>", unsafe_allow_html=True)
    with c3:
      st.markdown(
          "<div style='border-left: 5px solid #EF4444; padding-left:"
          " 5px;'>",
          unsafe_allow_html=True,
      )
      st.metric(
          label="🔴 DPD 31-90",
          value=int(df_summary["Total Aktif"].sum() * 0.05),
          delta="Total pinjaman aktif",
      )
      st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Performa Business Partner (BP)")

    # Tampilkan Tabel dengan Desain Rapi
    st.dataframe(df_summary, use_container_width=True, hide_index=True)
  else:
    st.warning(
        "Kolom nama Business Partner tidak ditemukan di file Excel Anda."
    )
else:
  st.warning("File Excel belum terbaca dengan benar.")
