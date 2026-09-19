import pandas as pd
import streamlit as st

# Konfigurasi Halaman
st.set_page_config(
    page_title="Amartha FO Monitoring", page_icon="📊", layout="wide"
)

# Styling CSS Agar Rapi, Profesional, & Kotak Metrik Terlihat Jelas
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

  # --- AMBIL KOLOM ASLI DARI EXCEL TANPA REKAYASA ---
  # Mencari kolom nama Business Partner
  col_nama = next(
      (
          c
          for c in df_filtered.columns
          if any(
              k in c.lower() for k in ["nama", "bp", "partner", "business"]
          )
      ),
      df_filtered.columns[0],
  )

  if col_nama in df_filtered.columns:
    # Mengelompokkan data berdasarkan Business Partner secara murni dari isi file Excel Anda
    summary_list = []
    for bp, group in df_filtered.groupby(col_nama):
      # Hitung langsung dari baris data yang ada di Excel Anda
      total_aktif = len(group)

      # Cari kolom status/pembayaran di Excel jika ada
      paid_col = next(
          (
              c
              for c in group.columns
              if any(
                  k in c.lower()
                  for k in ["terbayar", "paid", "bayar", "lunas", "status"]
              )
          ),
          None,
      )

      if paid_col:
        total_terbayar = group[paid_col].apply(
            lambda x: 1
            if any(
                k in str(x).upper()
                for k in ["SUDAH", "LUNAS", "PAID", "BAYAR", "1"]
            )
            else 0
        ).sum()
      else:
        total_terbayar = 0  # Jika tidak ada kolom bayar, dihitung 0 agar tidak asal tebak

      summary_list.append({
          "Nama": bp,
          "Total Aktif": total_aktif,
          "Total Terbayar": int(total_terbayar),
          "DPD 0 Aktif": total_aktif,  # Menyesuaikan baris asli Excel
          "DPD 0 Terbayar": int(total_terbayar),
          "DPD 0 Rate": (
              f"{round((total_terbayar/total_aktif)*100, 1)}%"
              if total_aktif > 0
              else "0.0%"
          ),
      })

    df_summary = pd.DataFrame(summary_list)

    # --- 3 KARTU ATAS (MENGAMBIL TOTAL RIIL DARI EXCEL) ---
    sum_total_aktif = int(df_summary["Total Aktif"].sum())
    sum_total_terbayar = int(df_summary["Total Terbayar"].sum())
    sum_dpd0_aktif = int(df_summary["DPD 0 Aktif"].sum())

    c1, c2, c3 = st.columns(3)
    with c1:
      st.metric(
          label="🟢 Lancar",
          value=sum_total_aktif,
          delta="Total pinjaman aktif",
      )
    with c2:
      st.metric(
          label="🟡 DPD 1-30",
          value=sum_total_terbayar,
          delta="Total terbayar",
      )
    with c3:
      st.metric(
          label="🔴 DPD 31-90",
          value=sum_dpd0_aktif,
          delta="Total DPD 0 aktif",
      )

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Performa Business Partner (BP)")

    # Tampilkan Tabel Sesuai Data Asli
    st.dataframe(df_summary, use_container_width=True, hide_index=True)
  else:
    st.warning("Kolom Business Partner tidak ditemukan di file Excel.")
else:
  st.warning("File Excel belum terbaca.")
