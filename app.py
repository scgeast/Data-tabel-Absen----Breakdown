import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

st.set_page_config(
    page_title="Aplikasi Proses Absensi",
    page_icon="📋",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 1.5rem;
        border-radius: 1rem;
        margin-bottom: 2rem;
        color: white;
        text-align: center;
    }
    .shift-badge {
        background-color: #dbeafe;
        padding: 0.2rem 0.5rem;
        border-radius: 1rem;
        display: inline-block;
        font-size: 0.8rem;
    }
    .summary-card {
        background: linear-gradient(135deg, #3b82f6, #1e40af);
        padding: 1rem;
        border-radius: 1rem;
        color: white;
        text-align: center;
    }
    .info-box {
        background-color: #fef3c7;
        padding: 0.8rem;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
        font-size: 0.8rem;
        border-left: 4px solid #f59e0b;
    }
    </style>
""", unsafe_allow_html=True)

# Shift mapping berdasarkan file
SHIFT_MAP = {
    ('DISPATCHER', 1): ('07:00:00', '15:00:00'),
    ('DISPATCHER', 2): ('15:00:00', '23:00:00'),
    ('DISPATCHER', 3): ('23:00:00', '06:00:00'),
    ('DISPATCHER', 11): ('07:00:00', '12:00:00'),
    ('DISPATCHER', 22): ('15:00:00', '20:00:00'),
    ('BOOKING EAST', 1): ('08:00:00', '16:00:00'),
    ('BOOKING EAST', 11): ('08:00:00', '13:00:00')
}

# Daftar hari dalam bahasa Indonesia
DAYS = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']

def get_day_name(date_obj):
    """Get Indonesian day name"""
    return DAYS[date_obj.weekday()]

def parse_shift_code(value):
    """Parse shift code from cell value"""
    if pd.isna(value):
        return None
    
    str_val = str(value).strip().upper()
    
    if str_val == 'OFF':
        return 'OFF'
    if str_val == 'CUTI':
        return 'CUTI'
    
    try:
        num = int(float(str_val))
        if num in [1, 2, 3, 11, 22]:
            return num
    except:
        pass
    
    return None

def process_excel(file_contents):
    """Process Excel file and return attendance data"""
    try:
        # Read Excel file
        excel_file = pd.ExcelFile(io.BytesIO(file_contents))
        
        # Read Sheet1
        if 'Sheet1' not in excel_file.sheet_names:
            st.error("Sheet 'Sheet1' tidak ditemukan dalam file")
            return None
        
        sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1', header=None)
        
        # Try to read INDEX sheet for shift mapping
        shift_map = SHIFT_MAP.copy()
        if 'INDEX' in excel_file.sheet_names:
            index_sheet = pd.read_excel(excel_file, sheet_name='INDEX', header=None)
            for _, row in index_sheet.iterrows():
                try:
                    if len(row) >= 5:
                        pos = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
                        num = int(row.iloc[2]) if pd.notna(row.iloc[2]) else None
                        check_in = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ''
                        check_out = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ''
                        
                        if pos in ['DISPATCHER', 'BOOKING EAST'] and num in [1, 2, 3, 11, 22]:
                            shift_map[(pos, num)] = (check_in, check_out)
                except:
                    continue
        
        # Generate date list: 11 May 2026 - 10 June 2026 (31 days)
        start_date = datetime(2026, 5, 11)
        date_list = [start_date + timedelta(days=i) for i in range(31)]
        
        # Find header row (contains "Jobdesk")
        header_row = -1
        for idx, row in sheet1.iterrows():
            if pd.notna(row.iloc[0]) and 'Jobdesk' in str(row.iloc[0]):
                header_row = idx
                break
        
        # Find data start row (first row with DISPATCHER or BOOKING EAST)
        data_start = -1
        for idx, row in sheet1.iterrows():
            if idx <= header_row:
                continue
            if pd.notna(row.iloc[0]):
                cell_val = str(row.iloc[0]).strip()
                if cell_val == 'DISPATCHER' or cell_val == 'BOOKING EAST':
                    data_start = idx
                    break
        
        if data_start == -1:
            st.error("Tidak menemukan data karyawan (DISPATCHER/BOOKING EAST)")
            return None
        
        # Process data
        results = []
        
        for idx in range(data_start, len(sheet1)):
            row = sheet1.iloc[idx]
            
            work_area = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            
            if work_area not in ['DISPATCHER', 'BOOKING EAST']:
                continue
            
            employee_name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
            if not employee_name or employee_name == 'Name' or employee_name == 'nan':
                continue
            
            # Process each date column (col 2 to col 33 = 31 days)
            for col_idx in range(2, min(34, len(row))):
                date_idx = col_idx - 2
                if date_idx >= len(date_list):
                    break
                
                current_date = date_list[date_idx]
                date_str = current_date.strftime('%d/%m/%Y')
                day_name = get_day_name(current_date)
                
                shift_value = row.iloc[col_idx]
                shift_code = parse_shift_code(shift_value)
                
                shift_label = ''
                check_in = ''
                check_out = ''
                remarks = ''
                
                if shift_code == 'OFF':
                    shift_label = 'OFF'
                    remarks = 'Libur / Off'
                elif shift_code == 'CUTI':
                    shift_label = 'CUTI'
                    remarks = 'Cuti'
                elif isinstance(shift_code, int):
                    shift_label = f'Shift {shift_code}'
                    key = (work_area, shift_code)
                    if key in shift_map:
                        check_in, check_out = shift_map[key]
                    if not check_in:
                        remarks = 'Jam shift tidak terdefinisi'
                else:
                    if pd.notna(shift_value) and str(shift_value).strip():
                        shift_label = '?'
                        remarks = f'Kode: {shift_value}'
                    else:
                        shift_label = '-'
                        remarks = 'Tidak ada data'
                
                results.append({
                    'work_area': work_area,
                    'employee_name': employee_name,
                    'job_position': work_area,
                    'day': day_name,
                    'date': date_str,
                    'shift': shift_label,
                    'check_in': check_in,
                    'check_out': check_out,
                    'remarks': remarks,
                    'shift_code': shift_code
                })
        
        if not results:
            st.error("Tidak ada data absensi yang ditemukan")
            return None
        
        return results
    
    except Exception as e:
        st.error(f"Error memproses file: {str(e)}")
        return None

def main():
    # Header
    st.markdown("""
        <div class="main-header">
            <h1>📋 Aplikasi Proses Absensi</h1>
            <p>DISPATCHER & BOOKING EAST | Periode 11 Mei - 10 Juni 2026</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Info box
    st.markdown("""
        <div class="info-box">
            ⚠️ <strong>Panduan:</strong> Upload file Excel dengan format yang sesuai (Sheet1 berisi data absensi, INDEX berisi mapping shift).
            File akan diproses otomatis. Data yang ditampilkan adalah periode <strong>11 Mei 2026 - 10 Juni 2026</strong> (31 hari).
        </div>
    """, unsafe_allow_html=True)
    
    # File uploader
    uploaded_file = st.file_uploader(
        "📂 Pilih file Excel",
        type=['xlsx', 'xls'],
        help="Upload file Absensi 11May'26~10June'26.xlsx"
    )
    
    if uploaded_file is not None:
        with st.spinner('Memproses file...'):
            file_contents = uploaded_file.read()
            data = process_excel(file_contents)
        
        if data:
            st.success(f"✅ Berhasil memproses {len(data)} catatan absensi!")
            
            # Convert to DataFrame
            df = pd.DataFrame(data)
            
            # Tabs
            tab1, tab2 = st.tabs(["📋 Data Absensi", "📊 Summary Shift"])
            
            with tab1:
                # Export button
                col1, col2, col3 = st.columns([1, 1, 4])
                with col1:
                    # Create Excel file for download
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_download = df[[
                            'work_area', 'employee_name', 'job_position', 
                            'day', 'date', 'shift', 'check_in', 'check_out', 'remarks'
                        ]].copy()
                        df_download.columns = [
                            'Work Area', 'Employee Name', 'Job Position',
                            'Day', 'Date', 'Shift', 'Check In', 'Check Out', 'Remarks'
                        ]
                        df_download.to_excel(writer, sheet_name='Absensi Data', index=False)
                    
                    st.download_button(
                        label="📎 Download Excel",
                        data=output.getvalue(),
                        file_name=f"hasil_absensi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Display dataframe
                st.dataframe(
                    df[[
                        'work_area', 'employee_name', 'job_position',
                        'day', 'date', 'shift', 'check_in', 'check_out', 'remarks'
                    ]].rename(columns={
                        'work_area': 'Work Area',
                        'employee_name': 'Employee Name',
                        'job_position': 'Job Position',
                        'day': 'Day',
                        'date': 'Date',
                        'shift': 'Shift',
                        'check_in': 'Check In',
                        'check_out': 'Check Out',
                        'remarks': 'Remarks'
                    }),
                    use_container_width=True,
                    height=500
                )
                
                st.caption(f"Total: {len(df)} baris data")
            
            with tab2:
                # Summary statistics
                col1, col2, col3, col4 = st.columns(4)
                
                unique_employees = df['employee_name'].nunique()
                unique_dates = df['date'].nunique()
                shift_aktif = df[df['shift_code'].apply(lambda x: isinstance(x, int))].shape[0]
                off_count = df[df['shift'] == 'OFF'].shape[0]
                
                with col1:
                    st.markdown(f"""
                        <div class="summary-card">
                            <h4>👥 Total Karyawan</h4>
                            <h2>{unique_employees}</h2>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                        <div class="summary-card" style="background: linear-gradient(135deg, #10b981, #065f46);">
                            <h4>📅 Total Hari Kerja</h4>
                            <h2>{unique_dates}</h2>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    st.markdown(f"""
                        <div class="summary-card" style="background: linear-gradient(135deg, #f59e0b, #b45309);">
                            <h4>🔄 Total Shift Aktif</h4>
                            <h2>{shift_aktif}</h2>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    st.markdown(f"""
                        <div class="summary-card" style="background: linear-gradient(135deg, #8b5cf6, #5b21b6);">
                            <h4>📆 Total Hari Off</h4>
                            <h2>{off_count}</h2>
                        </div>
                    """, unsafe_allow_html=True)
                
                # Summary per employee
                st.markdown("---")
                st.subheader("📊 Rekap Shift per Karyawan")
                
                # Calculate summary
                summary = {}
                for _, row in df.iterrows():
                    key = f"{row['employee_name']}|{row['job_position']}"
                    
                    if key not in summary:
                        summary[key] = {
                            'Employee Name': row['employee_name'],
                            'Jobdesk': row['job_position'],
                            '1 (Morning)': 0,
                            '2 (Afternoon)': 0,
                            '3 (Night)': 0,
                            '11 (Morning Part)': 0,
                            '22 (Afternoon Part)': 0,
                            'OFF': 0,
                            'CUTI': 0,
                            'TOTAL': 0
                        }
                    
                    code = row['shift_code']
                    if code == 'OFF':
                        summary[key]['OFF'] += 1
                    elif code == 'CUTI':
                        summary[key]['CUTI'] += 1
                    elif isinstance(code, int):
                        if code == 1:
                            summary[key]['1 (Morning)'] += 1
                        elif code == 2:
                            summary[key]['2 (Afternoon)'] += 1
                        elif code == 3:
                            summary[key]['3 (Night)'] += 1
                        elif code == 11:
                            summary[key]['11 (Morning Part)'] += 1
                        elif code == 22:
                            summary[key]['22 (Afternoon Part)'] += 1
                    
                    summary[key]['TOTAL'] += 1
                
                df_summary = pd.DataFrame(list(summary.values()))
                st.dataframe(df_summary, use_container_width=True, hide_index=True)
                
                # Shift information
                st.markdown("---")
                st.markdown("""
                    <div class="info-box">
                        ℹ️ <strong>Keterangan Shift:</strong><br>
                        • Shift 1 (Morning): 07:00 - 15:00<br>
                        • Shift 2 (Afternoon): 15:00 - 23:00<br>
                        • Shift 3 (Night): 23:00 - 06:00<br>
                        • Shift 11 (Morning Part): 07:00 - 12:00<br>
                        • Shift 22 (Afternoon Part): 15:00 - 20:00
                    </div>
                """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
