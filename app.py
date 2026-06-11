import streamlit as st
import pandas as pd
from datetime import datetime

# --- SHIFT MAPPING (dari INDEX) ---
SHIFT_MAP = {
    ("DISPATCHER", "1"): {"check_in": "07:00", "check_out": "15:00"},
    ("DISPATCHER", "2"): {"check_in": "15:00", "check_out": "23:00"},
    ("DISPATCHER", "3"): {"check_in": "23:00", "check_out": "06:00"},
    ("DISPATCHER", "11"): {"check_in": "07:00", "check_out": "12:00"},
    ("DISPATCHER", "22"): {"check_in": "15:00", "check_out": "20:00"},
    ("BOOKING EAST", "1"): {"check_in": "08:00", "check_out": "16:00"},
    ("BOOKING EAST", "11"): {"check_in": "08:00", "check_out": "13:00"},
}

def parse_excel_date(date_str):
    try:
        # Format: M/D/YY → e.g., "5/11/26"
        d = datetime.strptime(date_str, "%m/%d/%y")
        return d.strftime("%d/%m/%Y")
    except Exception:
        return str(date_str)

st.title("📄 Absensi → Format Laporan")
st.caption("Upload file Excel absensi (format seperti contoh Anda) untuk dikonversi ke tabel standar.")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None)

        # Cari baris dengan "Jobdesk" dan "Name" yang *berdekatan* dan *di bawah tanggal/hari*
        # Kita tahu: baris 10 = tanggal, baris 9 = hari, jadi data karyawan mulai baris 13+
        # Alternatif: cari baris yang kolom 0 == "DISPATCHER" atau "BOOKING EAST"
        data_rows = []
        for i in range(df.shape[0]):
            val0 = str(df.iloc[i, 0]).strip()
            if val0 in ["DISPATCHER", "BOOKING EAST"]:
                # Pastikan kolom 1 ada nama
                val1 = str(df.iloc[i, 1]).strip()
                if val1 and val1 not in ["nan", "Name"]:
                    data_rows.append(i)

        if not data_rows:
            st.error("❌ Tidak ditemukan baris karyawan (kolom 0 = 'DISPATCHER' / 'BOOKING EAST').")
        else:
            # Ambil baris tanggal (baris 10) dan hari (baris 9)
            date_row = df.iloc[10].dropna().tolist()
            day_row = df.iloc[9].dropna().tolist()

            # Filter hanya kolom yang berisi tanggal (format M/D/YY)
            dates = []
            days = []
            for i, cell in enumerate(date_row):
                if isinstance(cell, str) and "/" in cell:
                    dates.append(parse_excel_date(cell))
                    # Ambil hari dari baris 9 pada posisi yang sama
                    if i < len(day_row):
                        d = str(day_row[i]).strip()
                        if d in ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]:
                            days.append(d)
                        else:
                            # fallback: hitung dari tanggal
                            try:
                                dt = datetime.strptime(dates[-1], "%d/%m/%Y")
                                days.append(dt.strftime("%a").title())
                            except:
                                days.append("Unknown")
                    else:
                        days.append("Unknown")

            if not dates:
                st.error("❌ Tidak ditemukan kolom tanggal.")
            else:
                records = []
                for row_idx in data_rows:
                    jobdesk = str(df.iloc[row_idx, 0]).strip()
                    name = str(df.iloc[row_idx, 1]).strip()
                    if not name or name == "nan":
                        continue

                    # Ambil nilai absensi dari kolom 2 ke kanan (sesuai jumlah tanggal)
                    att_vals = df.iloc[row_idx, 2:2+len(dates)].astype(str).tolist()

                    for j, val in enumerate(att_vals):
                        val = val.strip()
                        if val in ["", "nan"]:
                            continue

                        remarks = ""
                        shift_code = val
                        if val in ["OFF", "CUTI"]:
                            remarks = val
                            shift_code = ""

                        # Cari mapping check-in/out berdasarkan (Work Area, Shift)
                        key = (jobdesk, shift_code)
                        ci = SHIFT_MAP.get(key, {}).get("check_in", "")
                        co = SHIFT_MAP.get(key, {}).get("check_out", "")

                        records.append({
                            "Work Area": jobdesk,
                            "Employee Name": name,
                            "Job Position": shift_code if shift_code else "",
                            "Day": days[j] if j < len(days) else "Unknown",
                            "Date (DD/MM/YYYY)": dates[j] if j < len(dates) else "",
                            "Shift": shift_code,
                            "Shift Check In": ci,
                            "Shift Check Out": co,
                            "Remarks": remarks
                        })

                if not records:
                    st.warning("⚠️ Tidak ada nilai absensi yang valid (semua sel kosong atau tidak terbaca).")
                else:
                    result_df = pd.DataFrame(records)
                    st.success(f"✅ Berhasil memproses {len(records)} baris data.")
                    st.dataframe(result_df)
                    csv = result_df.to_csv(index=False)
                    st.download_button("📥 Download CSV", csv, "absensi_formatted.csv", "text/csv")

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)
