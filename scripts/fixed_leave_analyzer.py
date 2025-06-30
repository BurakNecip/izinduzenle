import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch, cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

class FlexibleLeaveAnalyzer:
    def __init__(self):
        self.df = None
        self.column_mapping = {}
        self.setup_gui()
        # Register Turkish font for PDF
        self.setup_turkish_font()
    
    def setup_turkish_font(self):
        """Setup Turkish font support for PDF"""
        try:
            # Try to register a Turkish-compatible font
            # You can download DejaVuSans.ttf and put it in the same folder
            font_path = "DejaVuSans.ttf"
            if os.path.exists(font_path):
                pdfmetrics.registerFont(TTFont('DejaVuSans', font_path))
                self.turkish_font = 'DejaVuSans'
            else:
                # Fallback to Helvetica with proper encoding
                self.turkish_font = 'Helvetica'
        except:
            self.turkish_font = 'Helvetica'
    
    def setup_gui(self):
        self.root = tk.Tk()
        self.root.title("Esnek İzin Analiz Sistemi")
        self.root.geometry("800x700")
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Esnek İzin Analiz Sistemi", 
                               font=('Arial', 18, 'bold'))
        title_label.pack(pady=10)
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="Excel Dosyası Seçimi", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        file_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(file_frame, text="Dosya Seç", command=self.select_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="Verileri Yükle", command=self.load_data).pack(side=tk.LEFT, padx=5)
        
        # Column mapping frame
        self.mapping_frame = ttk.LabelFrame(main_frame, text="Sütun Eşleştirme", padding="10")
        self.mapping_frame.pack(fill=tk.X, pady=10)
        self.mapping_frame.pack_forget()  # Initially hidden
        
        # Date selection frame
        date_frame = ttk.LabelFrame(main_frame, text="Analiz Tarihi", padding="10")
        date_frame.pack(fill=tk.X, pady=10)
        
        # Start date
        start_frame = ttk.Frame(date_frame)
        start_frame.pack(fill=tk.X, pady=5)
        ttk.Label(start_frame, text="Başlangıç Tarihi (GG/AA/YYYY):").pack(side=tk.LEFT)
        self.start_date_var = tk.StringVar(value="21/07/2025")
        ttk.Entry(start_frame, textvariable=self.start_date_var, width=15).pack(side=tk.LEFT, padx=10)
        
        # End date
        end_frame = ttk.Frame(date_frame)
        end_frame.pack(fill=tk.X, pady=5)
        ttk.Label(end_frame, text="Bitiş Tarihi (GG/AA/YYYY):").pack(side=tk.LEFT)
        self.end_date_var = tk.StringVar(value="08/09/2025")
        ttk.Entry(end_frame, textvariable=self.end_date_var, width=15).pack(side=tk.LEFT, padx=10)
        
        # Generate report button
        report_frame = ttk.Frame(main_frame)
        report_frame.pack(pady=20)
        
        ttk.Button(report_frame, text="🔍 Analiz Yap", 
                  command=self.analyze_data, style='Accent.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(report_frame, text="📄 PDF Rapor Oluştur", 
                  command=self.generate_report).pack(side=tk.LEFT, padx=5)
        
        # Status and results area
        results_frame = ttk.LabelFrame(main_frame, text="Sonuçlar ve Durum", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Create text widget with scrollbar
        text_container = ttk.Frame(results_frame)
        text_container.pack(fill=tk.BOTH, expand=True)
        
        self.status_text = tk.Text(text_container, height=15, width=80, wrap=tk.WORD,
                                  font=('Consolas', 10))
        scrollbar = ttk.Scrollbar(text_container, orient=tk.VERTICAL, command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Status bar
        self.status_var = tk.StringVar(value="Hazır - Excel dosyası seçin")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, font=('Arial', 9))
        status_bar.pack(fill=tk.X, pady=5)
        
        # Store weekly data for report generation
        self.weekly_data = []
    
    def log(self, message):
        """Add message to status text"""
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update()
    
    def clear_log(self):
        """Clear status text"""
        self.status_text.delete(1.0, tk.END)
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Excel Dosyası Seçin",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.status_var.set(f"Dosya seçildi: {os.path.basename(file_path)}")
    
    def create_column_mapping_ui(self):
        """Create UI for column mapping"""
        # Clear existing widgets
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()
        
        ttk.Label(self.mapping_frame, text="Lütfen sütunları eşleştirin:", 
                 font=('Arial', 12, 'bold')).pack(pady=5)
        
        # Create mapping dropdowns
        self.column_vars = {}
        
        mappings = [
            ('name', 'İsim/Ad Sütunu:'),
            ('admin_start', 'İdari İzin Başlama:'),
            ('admin_end', 'İdari İzin Bitiş:'),
            ('annual_start', 'Yıllık İzin Başlama:'),
            ('annual_end', 'Yıllık İzin Bitiş:')
        ]
        
        column_options = ['Seçiniz...'] + list(self.df.columns)
        
        for key, label in mappings:
            frame = ttk.Frame(self.mapping_frame)
            frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(frame, text=label, width=20).pack(side=tk.LEFT, padx=5)
            
            var = tk.StringVar()
            combo = ttk.Combobox(frame, textvariable=var, values=column_options, 
                               state='readonly', width=30)
            combo.pack(side=tk.LEFT, padx=5)
            
            self.column_vars[key] = var
            
            # Auto-select if possible
            auto_selected = self.auto_select_column(key)
            if auto_selected:
                var.set(auto_selected)
        
        # Confirm button
        ttk.Button(self.mapping_frame, text="✅ Eşleştirmeyi Onayla", 
                  command=self.confirm_mapping).pack(pady=10)
        
        self.mapping_frame.pack(fill=tk.X, pady=10)
    
    def auto_select_column(self, key):
        """Auto-select column based on key"""
        keywords = {
            'name': ['isim', 'İSİM', 'ad', 'name', 'adi', 'adı', 'çalışan', 'personel'],
            'admin_start': ['idari', 'başlama', 'başlangıç'],
            'admin_end': ['idari', 'bitiş', 'bitim'],
            'annual_start': ['yillik', 'yıllık', 'başlama', 'başlangıç'],
            'annual_end': ['yillik', 'yıllık', 'bitiş', 'bitim']
        }
        
        for col in self.df.columns:
            col_lower = col.lower()
            if key in keywords:
                if key == 'admin_start':
                    if 'idari' in col_lower and ('başlama' in col_lower or 'başlangıç' in col_lower):
                        return col
                elif key == 'admin_end':
                    if 'idari' in col_lower and ('bitiş' in col_lower or 'bitim' in col_lower):
                        return col
                elif key == 'annual_start':
                    if ('yillik' in col_lower or 'yıllık' in col_lower) and ('başlama' in col_lower or 'başlangıç' in col_lower):
                        return col
                elif key == 'annual_end':
                    if ('yillik' in col_lower or 'yıllık' in col_lower) and ('bitiş' in col_lower or 'bitim' in col_lower):
                        return col
                elif key == 'name':
                    if any(keyword in col_lower for keyword in keywords[key]):
                        return col
        
        return None
    
    def confirm_mapping(self):
        """Confirm column mapping"""
        self.column_mapping = {}
        
        for key, var in self.column_vars.items():
            selected = var.get()
            if selected != 'Seçiniz...':
                self.column_mapping[key] = selected
        
        # Check if name column is selected
        if 'name' not in self.column_mapping:
            messagebox.showerror("Hata", "İsim sütunu seçilmesi zorunludur!")
            return
        
        self.log(f"\n✅ Sütun eşleştirmesi tamamlandı:")
        for key, col in self.column_mapping.items():
            self.log(f"  • {key}: {col}")
        
        self.mapping_frame.pack_forget()
        self.status_var.set("✅ Sütun eşleştirmesi tamamlandı - Analiz yapabilirsiniz")
    
    def load_data(self):
        if not self.file_path_var.get():
            messagebox.showerror("Hata", "Lütfen bir Excel dosyası seçin!")
            return
        
        try:
            self.clear_log()
            self.log("📂 Excel dosyası yükleniyor...")
            self.status_var.set("Veriler yükleniyor...")
            
            # Load Excel file
            self.df = pd.read_excel(self.file_path_var.get())
            
            self.log(f"✅ Toplam {len(self.df)} kayıt yüklendi")
            self.log("\n📋 Bulunan sütunlar:")
            for i, col in enumerate(self.df.columns, 1):
                self.log(f"  {i}. {col}")
            
            # Convert date columns
            date_columns = []
            for col in self.df.columns:
                if any(word in col.lower() for word in ['tarih', 'TARİH']):
                    date_columns.append(col)
                    self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
            
            self.log(f"\n📅 Tarih sütunları dönüştürüldü: {len(date_columns)} adet")
            for col in date_columns:
                self.log(f"  • {col}")
            
            self.log("\n✅ Veri yükleme tamamlandı!")
            self.log("👆 Şimdi sütunları eşleştirin...")
            
            # Show column mapping UI
            self.create_column_mapping_ui()
            
            self.status_var.set(f"✅ {len(self.df)} kayıt yüklendi - Sütunları eşleştirin")
            
        except Exception as e:
            error_msg = f"❌ Dosya yüklenirken hata: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Hata", error_msg)
            self.status_var.set("❌ Hata oluştu")
    
    def parse_date(self, date_str):
        """Parse date string in DD/MM/YYYY format"""
        try:
            return datetime.strptime(date_str.strip(), '%d/%m/%Y')
        except:
            return None
    
    def is_on_leave(self, employee_row, check_date):
        """Check if employee is on leave on a specific date"""
        # Check administrative leave
        if 'admin_start' in self.column_mapping and 'admin_end' in self.column_mapping:
            admin_start = employee_row[self.column_mapping['admin_start']]
            admin_end = employee_row[self.column_mapping['admin_end']]
            if pd.notna(admin_start) and pd.notna(admin_end):
                if admin_start <= check_date <= admin_end:
                    return True, "İdari İzin"
        
        # Check annual leave
        if 'annual_start' in self.column_mapping and 'annual_end' in self.column_mapping:
            annual_start = employee_row[self.column_mapping['annual_start']]
            annual_end = employee_row[self.column_mapping['annual_end']]
            if pd.notna(annual_start) and pd.notna(annual_end):
                if annual_start <= check_date <= annual_end:
                    return True, "Yıllık İzin"
        
        return False, "Çalışıyor"
    
    def get_week_start(self, date):
        """Get Monday of the week"""
        days_since_monday = date.weekday()
        return date - timedelta(days=days_since_monday)
    
    def analyze_data(self):
        """Analyze the data and show results"""
        if self.df is None:
            messagebox.showerror("Hata", "Önce veri yükleyin!")
            return
        
        if not self.column_mapping or 'name' not in self.column_mapping:
            messagebox.showerror("Hata", "Önce sütun eşleştirmesi yapın!")
            return
        
        try:
            # Parse dates
            start_date = self.parse_date(self.start_date_var.get())
            end_date = self.parse_date(self.end_date_var.get())
            
            if not start_date or not end_date:
                messagebox.showerror("Hata", "Geçerli tarih formatı: GG/AA/YYYY")
                return
            
            if start_date > end_date:
                messagebox.showerror("Hata", "Başlangıç tarihi bitiş tarihinden büyük olamaz!")
                return
            
            self.clear_log()
            self.log("🔍 HAFTALİK ÇALIŞAN ANALİZİ BAŞLADI")
            self.log("=" * 50)
            self.log(f"📅 Analiz Dönemi: {start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}")
            self.log(f"👥 Toplam Çalışan: {len(self.df)}")
            self.log(f"📋 İsim Sütunu: {self.column_mapping['name']}")
            self.log("=" * 50)
            
            self.status_var.set("🔍 Haftalık analiz yapılıyor...")
            
            # Generate weekly data
            self.weekly_data = self.generate_weekly_data(start_date, end_date)
            
            # Display results
            total_employees = len(self.df)
            working_counts = [len(w['working_employees']) for w in self.weekly_data]
            avg_working = sum(working_counts) / len(working_counts) if working_counts else 0
            max_working = max(working_counts) if working_counts else 0
            min_working = min(working_counts) if working_counts else 0
            
            self.log(f"\n📊 GENEL İSTATİSTİKLER:")
            self.log(f"  • Analiz edilen hafta sayısı: {len(self.weekly_data)}")
            self.log(f"  • Ortalama çalışan sayısı: {avg_working:.1f}")
            self.log(f"  • En fazla çalışan sayısı: {max_working}")
            self.log(f"  • En az çalışan sayısı: {min_working}")
            self.log(f"  • Ortalama yoğunluk: %{(avg_working/total_employees*100):.1f}")
            
            self.log(f"\n📋 HAFTALIK DETAYLAR:")
            self.log("-" * 50)
            
            for i, week_data in enumerate(self.weekly_data, 1):
                working_count = len(week_data['working_employees'])
                percentage = (working_count / total_employees * 100) if total_employees > 0 else 0
                
                # Status emoji
                if percentage >= 80:
                    status = "🟢 Yüksek"
                elif percentage >= 60:
                    status = "🟡 Orta"
                else:
                    status = "🔴 Düşük"
                
                self.log(f"\n{i}. {week_data['week_label']}")
                self.log(f"   Çalışan Sayısı: {working_count}/{total_employees}")
                self.log(f"   Yoğunluk: %{percentage:.1f} {status}")
                
                # Show first 10 employees
                if week_data['working_employees']:
                    self.log(f"   İlk 10 Çalışan:")
                    for j, emp in enumerate(week_data['working_employees'][:10], 1):
                        self.log(f"     {j:2d}. {emp}")
                    if len(week_data['working_employees']) > 10:
                        self.log(f"     ... ve {len(week_data['working_employees'])-10} kişi daha")
                else:
                    self.log(f"   ⚠️ Hiç çalışan yok!")
            
            self.log(f"\n✅ Analiz tamamlandı! PDF rapor oluşturabilirsiniz.")
            self.status_var.set("✅ Analiz tamamlandı - PDF rapor oluşturabilirsiniz")
            
        except Exception as e:
            error_msg = f"❌ Analiz sırasında hata: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Hata", error_msg)
            self.status_var.set("❌ Analiz hatası")
    
    def generate_weekly_data(self, start_date, end_date):
        """Generate weekly report data"""
        name_col = self.column_mapping['name']
        
        # Generate weeks
        current_date = self.get_week_start(start_date)
        weeks = []
        
        while current_date <= end_date:
            week_end = current_date + timedelta(days=6)
            weeks.append({
                'start': current_date,
                'end': week_end,
                'label': f"{current_date.strftime('%d %B')} Haftası"
            })
            current_date += timedelta(days=7)
        
        # Analyze each week
        weekly_reports = []
        
        for week in weeks:
            working_employees = []
            
            for index, employee in self.df.iterrows():
                employee_name = str(employee[name_col])
                
                # Check if working any day this week
                is_working = False
                for day_offset in range(5):  # Monday to Friday
                    check_date = week['start'] + timedelta(days=day_offset)
                    
                    if start_date <= check_date <= end_date:
                        on_leave, leave_type = self.is_on_leave(employee, check_date)
                        if not on_leave:
                            is_working = True
                            break
                
                if is_working:
                    working_employees.append(employee_name)
            
            weekly_reports.append({
                'week_label': week['label'],
                'working_employees': working_employees
            })
        
        return weekly_reports
    
    def create_modern_pdf_report(self, weekly_data, start_date, end_date, output_path):
        """Create modern PDF report with Turkish character support"""
        # Use landscape orientation for more space
        doc = SimpleDocTemplate(output_path, pagesize=landscape(A4), 
                              rightMargin=2*cm, leftMargin=2*cm, 
                              topMargin=1.5*cm, bottomMargin=1.5*cm)
        
        styles = getSampleStyleSheet()
        story = []
        
        # Custom styles with Turkish font support
        title_style = ParagraphStyle(
            'ModernTitle',
            parent=styles['Title'],
            fontSize=20,
            spaceAfter=20,
            alignment=1,  # Center
            textColor=colors.HexColor('#1976D2'),
            fontName=self.turkish_font
        )
        
        subtitle_style = ParagraphStyle(
            'ModernSubtitle',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=15,
            alignment=1,  # Center
            textColor=colors.HexColor('#424242'),
            fontName=self.turkish_font
        )
        
        week_header_style = ParagraphStyle(
            'WeekHeader',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=10,
            spaceBefore=15,
            textColor=colors.HexColor('#FF6F00'),
            fontName=self.turkish_font
        )
        
        normal_style = ParagraphStyle(
            'TurkishNormal',
            parent=styles['Normal'],
            fontName=self.turkish_font,
            fontSize=10
        )
        
        # Title page
        title = Paragraph("HAFTALİK ÇALIŞAN RAPORU", title_style)
        story.append(title)
        
        subtitle = Paragraph(
            f"Analiz Dönemi: {start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}", 
            subtitle_style
        )
        story.append(subtitle)
        
        # Report date
        report_date = Paragraph(
            f"Rapor Tarihi: {datetime.now().strftime('%d/%m/%Y %H:%M')}", 
            subtitle_style
        )
        story.append(report_date)
        story.append(Spacer(1, 30))
        
        # Summary statistics
        total_employees = len(self.df) if self.df is not None else 0
        working_counts = [len(w['working_employees']) for w in weekly_data]
        avg_working = sum(working_counts) / len(working_counts) if working_counts else 0
        max_working = max(working_counts) if working_counts else 0
        min_working = min(working_counts) if working_counts else 0
        
        # Summary table with Turkish characters
        summary_data = [
            ['ÖZET BİLGİLER', 'DEĞER'],
            ['Toplam Çalışan Sayısı', str(total_employees)],
            ['Analiz Edilen Hafta Sayısı', str(len(weekly_data))],
            ['Ortalama Çalışan Sayısı', f'{avg_working:.1f}'],
            ['En Fazla Çalışan Sayısı', str(max_working)],
            ['En Az Çalışan Sayısı', str(min_working)],
            ['Ortalama Çalışan Yoğunluğu', f'%{(avg_working/total_employees*100):.1f}' if total_employees > 0 else 'N/A']
        ]
        
        summary_table = Table(summary_data, colWidths=[6*cm, 4*cm])
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1976D2')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), self.turkish_font),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#E3F2FD')),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#1976D2')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))
        
        story.append(summary_table)
        story.append(PageBreak())  # New page for weekly details
        
        # Weekly reports
        for i, week_data in enumerate(weekly_data):
            # Week header
            week_header = Paragraph(f"{week_data['week_label']}", week_header_style)
            story.append(week_header)
            
            if week_data['working_employees']:
                # Create employee table
                employees = week_data['working_employees']
                
                # Split into multiple columns if too many employees
                employees_per_col = 30
                
                if len(employees) <= employees_per_col:
                    # Single column
                    table_data = [['#', 'ÇALIŞAN ADI']]
                    for j, employee in enumerate(employees, 1):
                        table_data.append([str(j), employee])
                    
                    employee_table = Table(table_data, colWidths=[1.5*cm, 10*cm])
                else:
                    # Multiple columns
                    col1 = employees[:employees_per_col]
                    col2 = employees[employees_per_col:employees_per_col*2] if len(employees) > employees_per_col else []
                    
                    table_data = [['#', 'ÇALIŞAN ADI', '#', 'ÇALIŞAN ADI']]
                    max_rows = max(len(col1), len(col2))
                    
                    for j in range(max_rows):
                        row = []
                        # Column 1
                        if j < len(col1):
                            row.extend([str(j+1), col1[j]])
                        else:
                            row.extend(['', ''])
                        
                        # Column 2
                        if j < len(col2):
                            row.extend([str(j+employees_per_col+1), col2[j]])
                        else:
                            row.extend(['', ''])
                        
                        table_data.append(row)
                    
                    employee_table = Table(table_data, colWidths=[1.5*cm, 8*cm, 1.5*cm, 8*cm])
                
                # Table styling with Turkish font
                employee_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FF6F00')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, -1), self.turkish_font),
                    ('FONTSIZE', (0, 0), (-1, 0), 11),
                    ('FONTSIZE', (0, 1), (-1, -1), 9),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E0E0E0')),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 8),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 8),
                    ('TOPPADDING', (0, 0), (-1, -1), 6),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ]))
                
                story.append(employee_table)
                story.append(Spacer(1, 15))
                
                # Week summary
                working_count = len(week_data['working_employees'])
                percentage = (working_count / total_employees * 100) if total_employees > 0 else 0
                
                summary_text = f"Bu hafta toplam {working_count} çalışan aktif görevde bulunmaktadır."
                if total_employees > 0:
                    summary_text += f" (Toplam çalışanların %{percentage:.1f}'i)"
                
                summary_para = Paragraph(summary_text, normal_style)
                story.append(summary_para)
            else:
                # No employees working
                no_employees = Paragraph("Bu hafta hiçbir çalışan aktif görevde bulunmamaktadır.", 
                                       normal_style)
                story.append(no_employees)
            
            # Add page break between weeks (except for the last one)
            if i < len(weekly_data) - 1:
                story.append(PageBreak())
        
        # Build PDF
        doc.build(story)
    
    def generate_report(self):
        if not self.weekly_data:
            messagebox.showerror("Hata", "Önce analiz yapın!")
            return
        
        try:
            # Parse dates
            start_date = self.parse_date(self.start_date_var.get())
            end_date = self.parse_date(self.end_date_var.get())
            
            if not start_date or not end_date:
                messagebox.showerror("Hata", "Geçerli tarih formatı: GG/AA/YYYY")
                return
            
            self.log("\n📄 PDF raporu oluşturuluyor...")
            self.status_var.set("📄 PDF raporu oluşturuluyor...")
            
            # Save PDF
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="PDF Raporu Kaydet"
            )
            
            if output_path:
                self.create_modern_pdf_report(self.weekly_data, start_date, end_date, output_path)
                self.log(f"✅ PDF raporu kaydedildi: {output_path}")
                messagebox.showinfo("Başarılı", f"PDF raporu oluşturuldu!\n{output_path}")
                self.status_var.set("✅ PDF raporu başarıyla oluşturuldu")
        
        except Exception as e:
            error_msg = f"❌ PDF raporu oluşturulurken hata: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Hata", error_msg)
            self.status_var.set("❌ PDF raporu hatası")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    print("Esnek İzin Analiz Sistemi başlatılıyor...")
    try:
        app = FlexibleLeaveAnalyzer()
        app.run()
    except Exception as e:
        print(f"Uygulama başlatılırken hata: {e}")
        input("Çıkmak için Enter'a basın...")
