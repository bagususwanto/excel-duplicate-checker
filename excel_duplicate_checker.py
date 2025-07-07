import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from pathlib import Path
import subprocess
import platform
from datetime import datetime
import json

class ExcelDuplicateChecker:
    def __init__(self, root):
        self.root = root
        self.current_file = None
        self.config_file = "duplicate_checker_config.json"
        
        # Load config first before setting up UI
        self.load_config()
        
        # Then setup UI
        self.setup_ui()
        
        # Bind event untuk menyimpan config saat aplikasi ditutup
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_ui(self):
        # Konfigurasi window utama
        self.root.title("Excel Duplicate Checker")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        self.root.configure(bg='#f0f0f0')
        
        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')
        
        # Custom styles
        style.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'), background='#f0f0f0')
        style.configure('Subtitle.TLabel', font=('Segoe UI', 10), background='#f0f0f0', foreground='#666')
        style.configure('Action.TButton', font=('Segoe UI', 11), padding=(20, 10))
        style.configure('Status.TLabel', font=('Segoe UI', 10), background='#f0f0f0')
        
        # Main container
        main_container = ttk.Frame(self.root, padding=30)
        main_container.pack(fill='both', expand=True)
        
        # Header section
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill='x', pady=(0, 30))
        
        # Title
        title_label = ttk.Label(header_frame, text="🔍 Excel Duplicate Checker", 
                               style='Title.TLabel')
        title_label.pack()
        
        # Subtitle
        subtitle_label = ttk.Label(header_frame, 
                                  text="Deteksi dan ekspor data duplikat dari file Excel dengan mudah",
                                  style='Subtitle.TLabel')
        subtitle_label.pack(pady=(5, 0))
        
        # Separator
        separator = ttk.Separator(main_container, orient='horizontal')
        separator.pack(fill='x', pady=(0, 30))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_container, text="📁 Pilih File Excel", padding=20)
        file_frame.pack(fill='x', pady=(0, 20))
        
        # File path display
        self.file_path_var = tk.StringVar()
        self.file_path_var.set("Belum ada file yang dipilih...")
        
        file_path_label = ttk.Label(file_frame, textvariable=self.file_path_var, 
                                   font=('Segoe UI', 9), foreground='#666')
        file_path_label.pack(fill='x', pady=(0, 10))
        
        # File selection button
        select_btn = ttk.Button(file_frame, text="📂 Pilih File Excel", 
                               command=self.pilih_file, style='Action.TButton')
        select_btn.pack()
        
        # Configuration section
        config_frame = ttk.LabelFrame(main_container, text="⚙️ Pengaturan", padding=20)
        config_frame.pack(fill='x', pady=(0, 20))
        
        # Skip rows setting
        skip_frame = ttk.Frame(config_frame)
        skip_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(skip_frame, text="Lewati baris dari atas:").pack(side='left')
        self.skip_rows_var = tk.StringVar(value=self.saved_skip_rows)
        skip_entry = ttk.Entry(skip_frame, textvariable=self.skip_rows_var, width=10)
        skip_entry.pack(side='right')
        
        # Column selection
        col_frame = ttk.Frame(config_frame)
        col_frame.pack(fill='x')
        
        ttk.Label(col_frame, text="Kolom untuk cek duplikat:").pack(side='left')
        self.column_var = tk.StringVar(value=self.saved_columns)
        col_entry = ttk.Entry(col_frame, textvariable=self.column_var, width=30)
        col_entry.pack(side='right')
        
        # Setelah membuat col_entry, tambahkan event binding:
        col_entry.bind('<KeyRelease>', self.on_column_change)
        col_entry.bind('<FocusOut>', self.on_column_change)
        skip_entry.bind('<KeyRelease>', self.on_skip_change)
        skip_entry.bind('<FocusOut>', self.on_skip_change)
        
        # Help text for column input
        help_label = ttk.Label(config_frame, text="💡 Pisahkan dengan koma untuk multiple kolom (contoh: Material No, Posting Date)", 
                              font=('Segoe UI', 8), foreground='#666')
        help_label.pack(pady=(5, 0))
        
        # Action section
        action_frame = ttk.LabelFrame(main_container, text="🚀 Proses", padding=20)
        action_frame.pack(fill='x', pady=(0, 20))
        
        # Process button
        self.process_btn = ttk.Button(action_frame, text="🔍 Cek Duplikat", 
                                     command=self.proses_duplikat, style='Action.TButton')
        self.process_btn.pack(pady=(0, 10))
        self.process_btn.configure(state='disabled')
        
        # Progress bar
        self.progress = ttk.Progressbar(action_frame, mode='indeterminate')
        self.progress.pack(fill='x', pady=(0, 10))
        
        # Status section
        status_frame = ttk.LabelFrame(main_container, text="📊 Status", padding=20)
        status_frame.pack(fill='both', expand=True)
        
        # Status text
        self.status_text = tk.Text(status_frame, height=6, wrap='word', 
                                  font=('Segoe UI', 9), state='disabled',
                                  bg='#f8f8f8', relief='flat', padx=10, pady=10)
        self.status_text.pack(fill='both', expand=True)
        
        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(status_frame, orient='vertical', command=self.status_text.yview)
        scrollbar.pack(side='right', fill='y')
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # Initial status
        self.update_status("Siap untuk memproses file Excel.", "info")
        
    def pilih_file(self):
        """Memilih file Excel untuk diproses"""
        file_path = filedialog.askopenfilename(
            title="Pilih File Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.current_file = file_path
            filename = os.path.basename(file_path)
            # Menampilkan path lengkap, bukan hanya filename
            self.file_path_var.set(f"📄 {file_path}")
            self.process_btn.configure(state='normal')
            self.update_status(f"File dipilih: {file_path}", "success")
        else:
            self.update_status("Tidak ada file yang dipilih.", "warning")
    
    def proses_duplikat(self):
        """Memproses file untuk mencari duplikat"""
        if not self.current_file:
            self.update_status("Silakan pilih file terlebih dahulu.", "error")
            return
        
        try:
            # Validate inputs
            try:
                skip_rows = int(self.skip_rows_var.get())
            except ValueError:
                skip_rows = 0
                
            column_input = self.column_var.get().strip()
            if not column_input:
                self.update_status("Nama kolom tidak boleh kosong.", "error")
                return
            
            # Parse multiple columns (comma-separated)
            column_names = [col.strip() for col in column_input.split(',') if col.strip()]
            if not column_names:
                self.update_status("Nama kolom tidak boleh kosong.", "error")
                return
            
            # Start processing
            self.update_status("Memproses file...", "info")
            self.progress.start()
            self.process_btn.configure(state='disabled')
            self.root.update_idletasks()
            
            # Read Excel file
            self.update_status(f"Membaca file Excel (melewati {skip_rows} baris)...", "info")
            df = pd.read_excel(self.current_file, skiprows=skip_rows)
            
            # Check if columns exist
            missing_cols = [col for col in column_names if col not in df.columns]
            if missing_cols:
                available_cols = ", ".join(df.columns.tolist()[:10])
                self.update_status(f"Kolom tidak ditemukan: {', '.join(missing_cols)}\nKolom tersedia: {available_cols}...", "error")
                return
            
            # Find duplicates
            column_display = ", ".join(column_names)
            self.update_status(f"Mencari duplikat berdasarkan kolom: {column_display}...", "info")
            duplikat = df[df.duplicated(subset=column_names, keep=False)]
            
            if not duplikat.empty:
                # Save duplicates
                file_path = Path(self.current_file)
                base_name = os.path.splitext(self.current_file)[0]
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                output_file = file_path.parent / f"{file_path.stem}_duplikat_{timestamp}.xlsx"
                output_file = str(output_file)
                duplikat.to_excel(output_file, index=False)
                
                # Simpan file bersih dari duplikat
                cleaned_df = df.drop_duplicates(subset=column_names, keep=False)
                clean_output_file = file_path.parent / f"{file_path.stem}_cleaned_{timestamp}.xlsx"
                clean_output_file = str(clean_output_file)
                cleaned_df.to_excel(clean_output_file, index=False)

                # Success message dengan path lengkap
                jumlah_duplikat = len(duplikat)
                jumlah_unik = len(duplikat.drop_duplicates(subset=column_names))
                
                success_msg = f"✅ DUPLIKAT DITEMUKAN!\n"
                success_msg += f"• Total baris duplikat: {jumlah_duplikat}\n"
                success_msg += f"• Jumlah set unik yang duplikat: {jumlah_unik}\n"
                success_msg += f"• Kolom yang dicek: {column_display}\n"
                success_msg += f"• File duplikat disimpan: {output_file}\n"
                success_msg += f"• File bersih disimpan: {clean_output_file}"
                
                self.update_status(success_msg, "success")
                
                # Show completion dialog dengan path lengkap
                if messagebox.askyesno("Proses Selesai", 
                                     f"Duplikat berhasil ditemukan!\n\n"
                                     f"Total baris duplikat: {jumlah_duplikat}\n"
                                     f"Kolom yang dicek: {column_display}\n\n"
                                     f"File duplikat disimpan di:\n{output_file}\n\n"
                                     f"File bersih disimpan di:\n{clean_output_file}\n\n"
                                     f"Apakah Anda ingin membuka folder output?"):
                    try:
                        # Buka folder output
                        output_folder = str(file_path.parent)
                        if platform.system() == "Windows":
                            os.startfile(output_folder)
                        elif platform.system() == "Darwin":  # macOS
                            subprocess.run(["open", output_folder])
                        else:  # Linux
                            subprocess.run(["xdg-open", output_folder])
                    except Exception as e:
                        self.update_status(f"Tidak dapat membuka folder: {str(e)}", "warning")
                    
            else:
                self.update_status(f"✅ TIDAK ADA DUPLIKAT\nData sudah bersih, tidak ada duplikat yang ditemukan.\nKolom yang dicek: {column_display}", "success")
                messagebox.showinfo("Hasil", f"Data tidak mengandung duplikat.\nKolom yang dicek: {column_display}")
                
        except Exception as e:
            error_msg = f"❌ KESALAHAN TERJADI\n{str(e)}"
            self.update_status(error_msg, "error")
            messagebox.showerror("Error", f"Terjadi kesalahan:\n{str(e)}")
            
        finally:
            # Reset UI
            self.progress.stop()
            self.process_btn.configure(state='normal')
    
    def update_status(self, message, status_type="info"):
        """Update status text dengan formatting"""
        self.status_text.configure(state='normal')
        self.status_text.delete(1.0, tk.END)
        
        # Add timestamp
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Color coding
        if status_type == "success":
            color = "#27ae60"
        elif status_type == "error":
            color = "#e74c3c"
        elif status_type == "warning":
            color = "#f39c12"
        else:
            color = "#2c3e50"
        
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.configure(state='disabled', foreground=color)
        self.status_text.see(tk.END)
        
    def load_config(self):
        """Load konfigurasi dari file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.saved_columns = config.get('columns', 'Material No')
                    self.saved_skip_rows = config.get('skip_rows', '6')
                print(f"Config loaded: {config}")  # Debug print
            else:
                self.saved_columns = 'Material No'
                self.saved_skip_rows = '6'
                print("No config file found, using defaults")  # Debug print
        except Exception as e:
            print(f"Error loading config: {e}")  # Debug print
            self.saved_columns = 'Material No'
            self.saved_skip_rows = '6'

    def on_column_change(self, event=None):
        """Handler untuk perubahan kolom"""
        self.save_config()
        
    def on_skip_change(self, event=None):
        """Handler untuk perubahan skip rows"""
        self.save_config()
        
    def save_config(self):
        """Simpan konfigurasi ke file"""
        try:
            config = {
                'columns': self.column_var.get(),
                'skip_rows': self.skip_rows_var.get()
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
            print(f"Config saved: {config}")  # Debug print
        except Exception as e:
            print(f"Error saving config: {e}")  # Debug print

    def on_closing(self):
        """Event handler saat aplikasi ditutup"""
        self.save_config()
        self.root.destroy()

def main():
    root = tk.Tk()
    app = ExcelDuplicateChecker(root)
    root.mainloop()

if __name__ == "__main__":
    main()