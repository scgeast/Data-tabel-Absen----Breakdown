import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

class AbsensiApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplikasi Proses Absensi - DISPATCHER & BOOKING EAST")
        self.root.geometry("1400x800")
        self.root.configure(bg='#f0f2f5')
        
        # Data
        self.attendance_data = None
        self.shift_map = {}
        
        self.setup_ui()
    
    def setup_ui(self):
        # Header
        header = tk.Frame(self.root, bg='white', height=80)
        header.pack(fill='x', padx=20, pady=(20,10))
        header.pack_propagate(False)
        
        tk.Label(header, text="📋 Aplikasi Proses Absensi", font=('Arial', 20, 'bold'), 
                 bg='white', fg='#1e293b').pack(side='left', padx=20)
        
        # Main frame
        main_frame = tk.Frame(self.root, bg='#f0f2f5')
        main_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Upload area
        upload_frame = tk.Frame(main_frame, bg='white', relief='raised', bd=1)
        upload_frame.pack(fill='x', pady=(0,20))
        
        tk.Button(upload_frame, text="📂 Pilih File Excel", command=self.load_file,
                  bg='#2563eb', fg='white', font=('Arial', 11, 'bold'),
                  padx=20, pady=10, cursor='hand2').pack(side='left', padx=20, pady=20)
        
        self.status_label = tk.Label(upload_frame, text="Belum ada file dipilih", 
                                      bg='white', fg='#64748b', font=('Arial', 10))
        self.status_label.pack(side='left', padx=20)
        
        # Notebook (tabs)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill='both', expand=True)
        
        # Tab Data
        self.data_frame = tk.Frame(self.notebook)
        self.notebook.add(self.data_frame, text="📋 Data Absensi")
        
        # Tab Summary
        self.summary_frame = tk.Frame(self.notebook)
        self.notebook.add(self.summary_frame, text="📊 Summary Shift")
        
        # Treeview untuk data
        self.setup_data_table()
        
        # Treeview untuk summary
        self.setup_summary_table()
        
        # Export button
        export_frame = tk.Frame(main_frame, bg='#f0f2f5')
        export_frame.pack(fill='x', pady=10)
        
        self.export_btn = tk.Button(export_frame, text="📎 Export to Excel", 
                                     command=self.export_to_excel,
                                     bg='#10b981', fg='white', font=('Arial', 11, 'bold'),
                                     padx=20, pady=8, state='disabled')
        self.export_btn.pack(side='right')
        
        # Info label
        info_label = tk.Label(main_frame, 
                              text="ℹ️ Shift: 1=Morning(07:00-15:00) | 2=Afternoon(15:00-23:00) | 3=Night(23:00-06:00) | 11=Morning Part(07:00-12:00) | 22=Afternoon Part(15:00-20:00)",
                              bg='#fef3c7', fg='#b45309', font=('Arial', 9), pady=5)
        info_label.pack(fill='x', pady=(10,0))
    
    def setup_data_table(self):
        # Frame dengan scrollbar
        frame = tk.Frame(self.data_frame)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Scrollbar
        vsb = tk.Scrollbar(frame, orient='vertical')
        hsb = tk.Scrollbar(frame, orient='horizontal')
        
        self.data_tree = ttk.Treeview(frame, 
                                       columns=('work_area', 'employee_name', 'job_position', 'day', 'date', 'shift', 'check_in', 'check_out', 'remarks'),
                                       show='headings',
                                       yscrollcommand=vsb.set,
                                       xscrollcommand=hsb.set)
        
        vsb.config(command=self.data_tree.yview)
        hsb.config(command=self.data_tree.xview)
        
        # Define columns
        columns = [
            ('work_area', 'Work Area', 120),
            ('employee_name', 'Employee Name', 130),
            ('job_position', 'Job Position', 120),
            ('day', 'Day', 100),
            ('date', 'Date', 110),
            ('shift', 'Shift', 100),
            ('check_in', 'Check In', 100),
            ('check_out', 'Check Out', 100),
            ('remarks', 'Remarks', 150)
        ]
        
        for col, text, width in columns:
            self.data_tree.heading(col, text=text, command=lambda c=col: self.sort_data(c))
            self.data_tree.column(col, width=width, anchor='center')
        
        self.data_tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        
        # Row count label
        self.row_count_label = tk.Label(self.data_frame, text="Total: 0 baris", 
                                         font=('Arial', 9), fg='#475569')
        self.row_count_label.pack(pady=5)
    
    def setup_summary_table(self):
        frame = tk.Frame(self.summary_frame)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        vsb = tk.Scrollbar(frame, orient='vertical')
        hsb = tk.Scrollbar(frame, orient='horizontal')
        
        self.summary_tree = ttk.Treeview(frame,
                                          columns=('name', 'jobdesk', 's1', 's2', 's3', 's11', 's22', 'off', 'cuti', 'total'),
                                          show='headings',
                                          yscrollcommand=vsb.set,
                                          xscrollcommand=hsb.set)
        
        vsb.config(command=self.summary_tree.yview)
        hsb.config(command=self.summary_tree.xview)
        
        summary_columns = [
            ('name', 'Employee Name', 150),
            ('jobdesk', 'Jobdesk', 120),
            ('s1', '1 (Morning)', 100),
            ('s2', '2 (Afternoon)', 100),
            ('s3', '3 (Night)', 80),
            ('s11', '11 (Morning Part)', 120),
            ('s22', '22 (Afternoon Part)', 130),
            ('off', 'OFF', 80),
            ('cuti', 'CUTI', 80),
            ('total', 'TOTAL', 80)
        ]
        
        for col, text, width in summary_columns:
            self.summary_tree.heading(col, text=text)
            self.summary_tree.column(col, width=width, anchor='center')
        
        self.summary_tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
    
    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Pilih file Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        self.status_label.config(text="⏳ Memproses file...", fg='#b45309')
        self.root.update()
        
        try:
            self.process_file(file_path)
            self.status_label.config(text=f"✅ Berhasil diproses: {os.path.basename(file_path)}", fg='#166534')
            self.export_btn.config(state='normal')
        except Exception as e:
            messagebox.showerror("Error", f"Gagal memproses file:\n{str(e)}")
            self.status_label.config(text="❌ Gagal memproses", fg='#991b1b')
    
    def build_shift_map(self, index_df):
        """Build shift mapping from INDEX sheet"""
        shift_map = {}
        
        # Default shifts based on the file
        default_shifts = {
            ('DISPATCHER', 1): ('07:00:00', '15:00:00'),
            ('DISPATCHER', 2): ('15:00:00', '23:00:00'),
            ('DISPATCHER', 3): ('23:00:00', '06:00:00'),
            ('DISPATCHER', 11): ('07:00:00', '12:00:00'),
            ('DISPATCHER', 22): ('15:00:00', '20:00:00'),
            ('BOOKING EAST', 1): ('08:00:00', '16:00:00'),
            ('BOOKING EAST', 11): ('08:00:00', '13:00:00')
        }
        
        # Try to read from INDEX sheet if available
        if index_df is not None and not index_df.empty:
            for _, row in index_df.iterrows():
                try:
                    pos = str(row.iloc[1]).strip() if len(row) > 1 else ''
                    num = int(row.iloc[2]) if len(row) > 2 and pd.notna(row.iloc[2]) else None
                    check_in = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ''
                    check_out = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else ''
                    
                    if pos in ['DISPATCHER', 'BOOKING EAST'] and num in [1,2,3,11,22]:
                        shift_map[(pos, num)] = (check_in, check_out)
                except:
                    continue
        
        # Fill missing with defaults
        for key, value in default_shifts.items():
            if key not in shift_map:
                shift_map[key] = value
        
        return shift_map
    
    def get_day_name(self, date_obj):
        """Get Indonesian day name"""
        days = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']
        return days[date_obj.weekday()]
    
    def parse_shift_code(self, value):
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
    
    def process_file(self, file_path):
        # Read Excel file
        excel_file = pd.ExcelFile(file_path)
        
        # Read sheets
        sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1', header=None)
        
        # Try to read INDEX sheet
        try:
            index_sheet = pd.read_excel(excel_file, sheet_name='INDEX', header=None)
        except:
            index_sheet = None
        
        # Build shift map
        self.shift_map = self.build_shift_map(index_sheet)
        
        # Generate date list: 11 May 2026 - 10 June 2026 (31 days)
        start_date = datetime(2026, 5, 11)
        date_list = [start_date + timedelta(days=i) for i in range(31)]
        
        # Find header row (contains "Jobdesk" or similar)
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
            raise ValueError("Tidak menemukan data karyawan (DISPATCHER/BOOKING EAST)")
        
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
                day_name = self.get_day_name(current_date)
                
                shift_value = row.iloc[col_idx]
                shift_code = self.parse_shift_code(shift_value)
                
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
                    if key in self.shift_map:
                        check_in, check_out = self.shift_map[key]
                    if not check_in:
                        remarks = 'Jam shift tidak terdefinisi'
                else:
                    if pd.notna(shift_value) and str(shift_value).strip():
                        shift_label = '?'
                        remarks = f'Kode tidak dikenal: {shift_value}'
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
            raise ValueError("Tidak ada data absensi yang ditemukan")
        
        self.attendance_data = results
        self.update_data_table()
        self.update_summary_table()
    
    def update_data_table(self):
        # Clear existing items
        for item in self.data_tree.get_children():
            self.data_tree.delete(item)
        
        # Insert data
        for row in self.attendance_data:
            self.data_tree.insert('', 'end', values=(
                row['work_area'],
                row['employee_name'],
                row['job_position'],
                row['day'],
                row['date'],
                row['shift'],
                row['check_in'],
                row['check_out'],
                row['remarks']
            ))
        
        self.row_count_label.config(text=f"Total: {len(self.attendance_data)} baris")
    
    def update_summary_table(self):
        # Calculate summary per employee
        summary = {}
        
        for row in self.attendance_data:
            key = f"{row['employee_name']}|{row['job_position']}"
            
            if key not in summary:
                summary[key] = {
                    'name': row['employee_name'],
                    'jobdesk': row['job_position'],
                    'shifts': {1: 0, 2: 0, 3: 0, 11: 0, 22: 0, 'OFF': 0, 'CUTI': 0},
                    'total': 0
                }
            
            code = row['shift_code']
            if code == 'OFF':
                summary[key]['shifts']['OFF'] += 1
            elif code == 'CUTI':
                summary[key]['shifts']['CUTI'] += 1
            elif isinstance(code, int):
                summary[key]['shifts'][code] += 1
            
            summary[key]['total'] += 1
        
        # Clear existing items
        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)
        
        # Insert summary data
        for key, data in summary.items():
            self.summary_tree.insert('', 'end', values=(
                data['name'],
                data['jobdesk'],
                data['shifts'][1],
                data['shifts'][2],
                data['shifts'][3],
                data['shifts'][11],
                data['shifts'][22],
                data['shifts']['OFF'],
                data['shifts']['CUTI'],
                data['total']
            ))
    
    def sort_data(self, column):
        # Map column to key
        col_map = {
            'work_area': 'work_area',
            'employee_name': 'employee_name',
            'job_position': 'job_position',
            'day': 'day',
            'date': 'date',
            'shift': 'shift',
            'check_in': 'check_in',
            'check_out': 'check_out',
            'remarks': 'remarks'
        }
        
        key = col_map.get(column, 'date')
        
        # Sort data
        self.attendance_data.sort(key=lambda x: x[key] if x[key] else '')
        
        # Update table
        self.update_data_table()
    
    def export_to_excel(self):
        if not self.attendance_data:
            messagebox.showwarning("Warning", "Tidak ada data untuk diekspor")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"hasil_absensi_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        
        if not file_path:
            return
        
        try:
            # Create DataFrames
            df_data = pd.DataFrame(self.attendance_data)
            df_data = df_data[[
                'work_area', 'employee_name', 'job_position', 'day', 
                'date', 'shift', 'check_in', 'check_out', 'remarks'
            ]]
            
            # Rename columns
            df_data.columns = [
                'Work Area', 'Employee Name', 'Job Position', 'Day',
                'Date', 'Shift', 'Check In', 'Check Out', 'Remarks'
            ]
            
            # Create summary DataFrame
            summary = {}
            for row in self.attendance_data:
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
                    col_name = f'{code} (Morning)' if code == 1 else f'{code} (Afternoon)' if code == 2 else f'{code} (Night)' if code == 3 else f'{code} (Morning Part)' if code == 11 else f'{code} (Afternoon Part)'
                    summary[key][col_name] += 1
                
                summary[key]['TOTAL'] += 1
            
            df_summary = pd.DataFrame(list(summary.values()))
            
            # Export to Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df_data.to_excel(writer, sheet_name='Absensi Data', index=False)
                df_summary.to_excel(writer, sheet_name='Summary Shift', index=False)
            
            messagebox.showinfo("Sukses", f"File berhasil diekspor ke:\n{file_path}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengekspor file:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AbsensiApp(root)
    root.mainloop()
