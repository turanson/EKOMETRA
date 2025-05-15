import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry  # Tarih seçimi için
import pandas as pd
import warnings

# OpenPyXL uyarılarını bastır
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


class TrendyolSiparisAnaliz:
    def __init__(self, root):
        self.root = root
        self.df = None  # Orijinal veri
        self.filtered_df = None  # Filtrelenmiş veri
        self.sort_order = True  # Sıralama düzeni (True: Artan, False: Azalan)
        self.setup_styles()  # Stil ayarlarını yap
        self.setup_ui()

    def setup_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("default")
        style.configure("Treeview", background="white", foreground="black", rowheight=25, fieldbackground="white")
        style.map("Treeview", background=[("selected", "#f0ad4e")])  # Seçili hücre rengi
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"), background="#f7f7f9", foreground="black", borderwidth=1)

    def setup_ui(self):
        # Üst başlık
        header = ttk.Label(self.root, text="Trendyol Sipariş Analiz ve Raporlama", font=("Arial", 18, "bold"))
        header.pack(pady=10)

        # Tarih filtreleme alanı
        filter_frame = ttk.Frame(self.root)
        filter_frame.pack(pady=5)

        ttk.Label(filter_frame, text="Başlangıç Tarihi:", font=("Arial", 10)).grid(row=0, column=0, padx=5, pady=5)
        self.start_date = DateEntry(filter_frame, width=12, background='darkblue', foreground='white', date_pattern='yyyy-mm-dd')
        self.start_date.grid(row=0, column=1, padx=5)

        ttk.Label(filter_frame, text="Bitiş Tarihi:", font=("Arial", 10)).grid(row=0, column=2, padx=5, pady=5)
        self.end_date = DateEntry(filter_frame, width=12, background='darkblue', foreground='white', date_pattern='yyyy-mm-dd')
        self.end_date.grid(row=0, column=3, padx=5)

        btn_filter = ttk.Button(filter_frame, text="Filtrele", command=self.apply_date_filter)
        btn_filter.grid(row=0, column=4, padx=10)

        # Dosya seçme butonu
        btn_yukle = ttk.Button(self.root, text="Excel Dosyası Seç", command=self.dosya_sec, width=30)
        btn_yukle.pack(pady=10)

        self.dosya_adi_label = ttk.Label(self.root, text="", foreground="blue", font=("Arial", 10))
        self.dosya_adi_label.pack(pady=5)

        # Genel toplam alanı
        self.genel_toplam_label = ttk.Label(
            self.root,
            text="Genel Toplam Adet: 0 | Genel Toplam Ciro: 0 TL",
            font=("Arial", 14, "bold"),
            foreground="green",
        )
        self.genel_toplam_label.pack(pady=10)

        # Rapor tablosu
        self.rapor_tree = ttk.Treeview(self.root, columns=("Kategori", "Değer"), show="headings", height=15)
        self.rapor_tree.heading("Kategori", text="Ürün Adı")
        self.rapor_tree.heading("Değer", text="Toplam Adet", command=self.sort_by_adet)
        self.rapor_tree.column("Kategori", width=300, anchor="center")
        self.rapor_tree.column("Değer", width=150, anchor="center")
        self.rapor_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Zebra deseni
        self.rapor_tree.tag_configure("evenrow", background="#f9f9f9")
        self.rapor_tree.tag_configure("oddrow", background="#ffffff")

        # Ürün seçimi için olay bağlama
        self.rapor_tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        # Detay tablosu
        self.detay_tree = ttk.Treeview(self.root, columns=("Bilgi", "Değer"), show="headings", height=5)
        self.detay_tree.heading("Bilgi", text="Bilgi")
        self.detay_tree.heading("Değer", text="Değer")
        self.detay_tree.column("Bilgi", width=300, anchor="center")
        self.detay_tree.column("Değer", width=150, anchor="center")
        self.detay_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def dosya_sec(self):
        dosya_yolu = filedialog.askopenfilename(filetypes=[("Excel Dosyaları", "*.xlsx *.xls")])
        if dosya_yolu:
            self.dosya_adi_label.config(text=f"Yüklenen Dosya: {dosya_yolu}")
            self.analiz_et(dosya_yolu)

    def analiz_et(self, dosya_yolu):
        try:
            self.df = pd.read_excel(dosya_yolu)

            required_columns = ['Sipariş Tarihi', 'Ürün Adı', 'Satış Tutarı', 'Adet']
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            if missing_columns:
                raise ValueError(f"Excel dosyasında eksik sütunlar var: {', '.join(missing_columns)}")

            self.df['Sipariş Tarihi'] = pd.to_datetime(self.df['Sipariş Tarihi'], dayfirst=True, errors='coerce')
            self.df = self.df.dropna(subset=['Sipariş Tarihi'])  # Geçersiz tarihleri kaldır

            self.filtered_df = self.df.copy()  # Başlangıçta tüm veri yüklü
            self.update_ui()
        except Exception as e:
            messagebox.showerror("Hata", f"İşlem sırasında bir hata oluştu:\n{str(e)}")

    def update_ui(self):
        self.apply_date_filter()  # Filtreyi uygula

    def apply_date_filter(self):
        try:
            start_date = pd.to_datetime(self.start_date.get())
            end_date = pd.to_datetime(self.end_date.get())
            self.filtered_df = self.df[(self.df['Sipariş Tarihi'] >= start_date) & (self.df['Sipariş Tarihi'] <= end_date)]

            grouped_data = self.filtered_df.groupby('Ürün Adı')['Adet'].sum().reset_index()
            self.populate_treeview(grouped_data)

            genel_toplam_adet = self.filtered_df['Adet'].sum()
            genel_toplam_ciro = self.filtered_df['Satış Tutarı'].sum()

            self.genel_toplam_label.config(
                text=f"Genel Toplam Adet: {genel_toplam_adet} | Genel Toplam Ciro: {genel_toplam_ciro:,.2f} TL"
            )
        except Exception as e:
            messagebox.showerror("Hata", f"Filtreleme sırasında bir hata oluştu:\n{str(e)}")

    def populate_treeview(self, data):
        self.rapor_tree.delete(*self.rapor_tree.get_children())
        for i, row in data.iterrows():
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.rapor_tree.insert("", "end", values=(row['Ürün Adı'], row['Adet']), tags=(tag,))

    def sort_by_adet(self):
        self.sort_order = not self.sort_order
        sorted_data = self.filtered_df.groupby('Ürün Adı')['Adet'].sum().reset_index().sort_values(by='Adet', ascending=self.sort_order)
        self.populate_treeview(sorted_data)

    def on_tree_select(self, event):
        selected_item = self.rapor_tree.focus()
        if not selected_item:
            return

        selected_values = self.rapor_tree.item(selected_item, "values")
        if len(selected_values) < 1:  # Ürün adı boşsa
            return

        selected_product = selected_values[0]  # Ürün adı
        detay_df = self.filtered_df[self.filtered_df['Ürün Adı'] == selected_product]

        # Toplam sipariş adedi, ortalama satış tutarı ve toplam ciroyu hesapla
        toplam_adet = detay_df['Adet'].sum()
        ortalama_satis_tutari = detay_df['Satış Tutarı'].mean()
        toplam_ciro = detay_df['Satış Tutarı'].sum()

        # Detay tablosunu doldur
        self.detay_tree.delete(*self.detay_tree.get_children())
        self.detay_tree.insert("", "end", values=("Toplam Sipariş Adedi", toplam_adet))
        self.detay_tree.insert("", "end", values=("Ortalama Satış Tutarı (TL)", f"{ortalama_satis_tutari:,.2f}"))
        self.detay_tree.insert("", "end", values=("Toplam Ciro (TL)", f"{toplam_ciro:,.2f}"))


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Trendyol Sipariş Analiz Uygulaması")
    root.geometry("800x600")
    app = TrendyolSiparisAnaliz(root)
    root.mainloop()