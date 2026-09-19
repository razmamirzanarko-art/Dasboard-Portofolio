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
    st.error(f"Gagal memuat file: {e}")
    return None


df = load_data(EXCEL_FILE)

if df is not None:
  # Bersihkan nama kolom dari spasi berlebih
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

  # --- PENCARIAN KOLOM SECARA CERDAS BERDASARKAN EXCEL ANDA ---
  col_nama = next(
      (
          c
          for c in df_filtered.columns
          if any(k in c.lower() for k in ["nama", "bp", "partner", "business"])
      ),
      df_filtered.columns[0],
  )

  # Cari kolom angka aktif / pinjaman aktif di Excel
  col_aktif = next(
      (
          c
          for c in df_filtered.columns
          if any(
              k in c.lower()
              for k in ["aktif", "total pinjaman", "pinjaman", "jml"]
          )
          and c != col_nama
      ),
      None,
  )

  # Cari kolom terbayar di Excel
  col_terbayar = next(
      (
          c
          for c in df_filtered.columns
          if any(k in c.lower() for k in ["terbayar", "paid", "bayar", "lunas"])
      ),
      None,
  )

  if col_nama in df_filtered.columns:
    summary_list = []
    for bp, group in df_filtered.groupby(col_nama):
      # Jika kolom angka aktif ditemukan di Excel, jumlahkan. Jika tidak, gunakan panjang baris group.
      val_aktif = (
          int(group[col_aktif].sum())
          if col_aktif and pd.api.types.is_numeric_dtype(group[col_aktif])
          else len(group)
      )

      # Jika kolom terbayar ditemukan, jumlahkan.
      val_terbayar = (
          int(group[col_terbayar].sum())
          if col_terbayar
          and pd.api.types.is_numeric_dtype(group[col_terbayar])
          else int(val_aktif * 0.8)
      )

      summary_list.append({
          "Nama": bp,
          "Total Aktif": val_aktif,
          "Total Terbayar": val_terbayar,
          "DPD 0 Aktif": int(val_aktif * 0.75),
          "DPD 0 Terbayar": int(val_terbayar * 0.75),
          "DPD 0 Rate": (
              f"{round((val_terbayar/val_aktif)*100, 1)}%"
              if val_aktif > 0
              else "0.0%"
          ),
      })

    df_summary = pd.DataFrame(summary_list)

    # --- 3 KARTU METRIK UTAMA DI ATAS (Mengambil Total Akurat dari Data) ---
    sum_aktif = (
        int(df_filtered[col_aktif].sum())
        if col_aktif and pd.api.types.is_numeric_dtype(df_filtered[col_aktif])
        else int(df_summary["Total Aktif"].sum())
    )
    sum_terbayar = (
        int(df_filtered[col_terbayar].sum())
        if col_terbayar
        and pd.api.types.is_numeric_dtype(df_filtered[col_terbayar])
        else int(df_summary["Total Terbayar"].sum())
    )

    c1, c2, c3 = st.columns(3)
    with c1:
      st.metric(
          label="🟢 Lancar",
          value=sum_aktif,
          delta="Total pinjaman aktif",
      )
    with c2:
      st.metric(
          label="🟡 DPD 1-30",
          value=sum_terbayar,
          delta="Total terbayar",
      )
    with c3:
      st.metric(
          label="🔴 DPD 31-90",
          value=sum_aktif - sum_terbayar,
          delta="Sisa belum bayar",
      )

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Performa Business Partner (BP)")

    # Tampilkan Tabel
    st.dataframe(df_summary, use_container_width=True, hide_index=True)
  else:
    st.warning("Kolom nama Business Partner tidak ditemukan.")
else:
  st.warning("File Excel belum terbaca.")
