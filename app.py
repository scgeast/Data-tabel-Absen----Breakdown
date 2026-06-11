import streamlit as st
import pandas as pd
from datetime import datetime

# --- Helper: Parse date from Excel header (e.g., "5/11/26") ---
def parse_excel_date(date_str):
    # Format: M/D/YY → convert to DD/MM/YYYY
    try:
        d = datetime.strptime(date_str, "%m/%d/%y")
        return d.strftime("%d/%m/%Y")
    except:
        return date_str

# --- Shift mapping from INDEX table ---
SHIFT_MAP = {
    "1": {"check_in": "07:00", "check_out": "15:00"},
    "2": {"check_in": "15:00", "check_out": "23:00"},
    "3": {"check_in": "23:00", "check_out": "06:00"},
    "11": {"check_in": "07:00", "check_out": "12:00"},
    "22": {"check_in": "15:00", "check_out": "20:00"},
}

# --- Main App ---
st.title("📄 Absensi → Format Laporan")
st.caption("Upload file Excel absensi (format seperti contoh Anda) untuk dikonversi ke tabel standar.")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None)

        # Find start of data: look for row with "Name" and "Jobdesk"
        name_row_idx = None
        for i, row in df.iterrows():
            if "Name" in str(row.values) and "Jobdesk" in str(row.values):
                name_row_idx = i
                break

        if name_row_idx is None:
            st.error("Tidak ditemukan header 'Name' dan 'Jobdesk'. Pastikan format file sesuai.")
        else:
            # Extract column headers (dates)
            date_headers = df.iloc[name_row_idx + 1].dropna().tolist()
            # Clean up: remove empty or non-date strings
            date_headers = [h for h in date_headers if isinstance(h, str) and "/" in h]
            dates = [parse_excel_date(h) for h in date_headers]

            # Extract day names (Mon, Tue, ...) — assume they're in row above dates (row name_row_idx)
            day_headers = df.iloc[name_row_idx].dropna().tolist()
            days = [d for d in day_headers if d in ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]]

            # Ensure alignment: should have same length
            if len(dates) != len(days):
                # Fallback: generate days from dates
                days = []
                for d_str in dates:
                    try:
                        d = datetime.strptime(d_str, "%d/%m/%Y")
                        days.append(d.strftime("%a").title())  # Mon, Tue, etc.
                    except:
                        days.append("Unknown")

            # Extract employee rows
            data_start_row = name_row_idx + 2
            employees = []
            for i in range(data_start_row, df.shape[0]):
                row = df.iloc[i]
                jobdesk = str(row.iloc[0]).strip()
                if not jobdesk or jobdesk == "nan":
                    continue
                name = str(row.iloc[1]).strip()
                if name in ["nan", ""]:
                    continue

                # Get attendance values (skip first 2 columns: Jobdesk, Name)
                att_values = row.iloc[2:].tolist()[:len(dates)]

                for j, val in enumerate(att_values):
                    val = str(val).strip()
                    if val in ["", "nan"]:
                        continue

                    # Determine remarks
                    remarks = ""
                    shift_code = val
                    if val in ["OFF", "CUTI"]:
                        remarks = val
                        shift_code = ""
                    else:
                        # Try to map shift code
                        if val not in SHIFT_MAP:
                            st.warning(f"Kode shift '{val}' tidak dikenali untuk {name} pada {dates[j]}")

                    # Build record
                    record = {
                        "Work Area": jobdesk,
                        "Employee Name": name,
                        "Job Position": shift_code if shift_code else "",
                        "Day": days[j],
                        "Date (DD/MM/YYYY)": dates[j],
                        "Shift": shift_code,
                        "Shift Check In": SHIFT_MAP.get(shift_code, {}).get("check_in", "") if shift_code else "",
                        "Shift Check Out": SHIFT_MAP.get(shift_code, {}).get("check_out", "") if shift_code else "",
                        "Remarks": remarks,
                    }
                    employees.append(record)

            # Convert to DataFrame
            result_df = pd.DataFrame(employees)
            st.success(f"✅ Berhasil memproses {len(employees)} baris data.")

            # Show preview
            st.dataframe(result_df.head(50))

            # Download button
            csv = result_df.to_csv(index=False)
            st.download_button(
                label="📥 Download CSV",
                data=csv,
                file_name="absensi_formatted.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"Error saat membaca file: {e}")
        st.exception(e)
