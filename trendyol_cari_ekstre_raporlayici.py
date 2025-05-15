import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os


class TrendyolCariEkstreRaporlayici:
    def __init__(self, root):
        self.root = root
        self.df = None  # Veri çerçevesi burada saklanacak
        self.setup_ui()

    def setup_ui(self):
        # Tema ve üst başlık
        style = ttk.Style()
        style.theme_use("clam")
        header = ttk.Label(self.root, text="Trendyol Cari Ekstre Raporlayıcı", font=("Arial", 18, "bold"))
        header.pack(pady=10)

        # Dosya seçme butonu
        lbl = ttk.Label(self.root, text="Excel dosyasını yükleyin:", font=("Arial", 12))
        lbl.pack(pady=5)

        btn_yukle = ttk.Button(self.root, text="Excel Dosyası Seç", command=self.dosya_sec, width=30)
        btn_yukle.pack(pady=10)

        self.dosya_adi_label = ttk.Label(self.root, text="", foreground="blue", font=("Arial", 10))
        self.dosya_adi_label.pack(pady=5)

        # Tablo ve grafik alanları
        self.rapor_frame = ttk.Frame(self.root)
        self.rapor_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Rapor metni
        self.rapor_text = tk.Text(self.rapor_frame, wrap=tk.WORD, font=("Consolas", 10), height=15)
        self.rapor_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Grafik alanı
        self.grafik_frame = ttk.Frame(self.root)
        self.grafik_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Excel'e kaydet butonu
        self.export_button = ttk.Button(self.root, text="Raporu Excel Olarak İndir", command=self.export_to_excel, state="disabled")
        self.export_button.pack(pady=10)

    def dosya_sec(self):
        dosya_yolu = filedialog.askopenfilename(filetypes=[("Excel Dosyaları", "*.xlsx *.xls")])
        if dosya_yolu:
            self.dosya_adi_label.config(text=os.path.basename(dosya_yolu))
            self.analiz_et(dosya_yolu)

    def analiz_et(self, dosya_yolu):
        try:
            # Excel dosyasını yükle
            df = pd.read_excel(dosya_yolu)

            # Beklenen sütunların varlığını kontrol et
            required_columns = ['Borç', 'Alacak', 'Fiş Türü']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Excel dosyasında eksik sütunlar var: {', '.join(missing_columns)}")

            # Veri temizleme
            df['Borç'] = pd.to_numeric(df['Borç'], errors='coerce').fillna(0)
            df['Alacak'] = pd.to_numeric(df['Alacak'], errors='coerce').fillna(0)

            toplam_borc = df['Borç'].sum()
            toplam_alacak = df['Alacak'].sum()
            net_tutar = toplam_alacak - toplam_borc

            # Gider ve gelir analizleri
            gider_analizi = df[df['Borç'] > 0].groupby('Fiş Türü')['Borç'].sum().reset_index()
            gelir_analizi = df[df['Alacak'] > 0].groupby('Fiş Türü')['Alacak'].sum().reset_index()

            # Raporu göster
            self.rapor_text.delete(1.0, tk.END)
            self.rapor_text.insert(tk.END, f"🔎 GENEL RAPOR\n")
            self.rapor_text.insert(tk.END, f"Toplam Alacak : {toplam_alacak:,.2f} TL\n")
            self.rapor_text.insert(tk.END, f"Toplam Borç   : {toplam_borc:,.2f} TL\n")
            self.rapor_text.insert(tk.END, f"Net Kazanç    : {net_tutar:,.2f} TL\n\n")

            self.rapor_text.insert(tk.END, "📊 GİDER ANALİZİ (Fiş Türü Bazlı)\n")
            for _, row in gider_analizi.iterrows():
                self.rapor_text.insert(tk.END, f"- {row['Fiş Türü']}: {row['Borç']:,.2f} TL\n")

            self.rapor_text.insert(tk.END, f"\n💰 GELİR ANALİZİ (Fiş Türü Bazlı)\n")
            for _, row in gelir_analizi.iterrows():
                self.rapor_text.insert(tk.END, f"- {row['Fiş Türü']}: {row['Alacak']:,.2f} TL\n")

            # Grafik gösterimi
            self.grafik_goster(gider_analizi, gelir_analizi)

            # Export butonunu aktif et
            self.export_button.config(state="normal")

            # Kaydedilecek veriyi sakla
            self.mismatched_data = {"gider_analizi": gider_analizi, "gelir_analizi": gelir_analizi}

        except ValueError as ve:
            messagebox.showwarning("Veri Hatası", str(ve))
        except Exception as e:
            messagebox.showerror("Hata", f"İşlem sırasında bir hata oluştu:\n{str(e)}")

    def grafik_goster(self, gider_analizi, gelir_analizi):
        # Eski grafikleri temizle
        for widget in self.grafik_frame.winfo_children():
            widget.destroy()

        # Gider grafiği
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.bar(gider_analizi['Fiş Türü'], gider_analizi['Borç'], color='tomato', label="Gider")
        ax.set_title("Gider Analizi", fontsize=14)
        ax.set_ylabel("Tutar (TL)")
        ax.legend()

        # Gelir grafiği
        ax2 = ax.twinx()
        ax2.bar(gelir_analizi['Fiş Türü'], gelir_analizi['Alacak'], color='limegreen', label="Gelir")
        ax2.legend(loc="upper right")

        # X-eksenindeki yazıları döndür
        ax.set_xticks(range(len(gider_analizi['Fiş Türü'])))
        ax.set_xticklabels(gider_analizi['Fiş Türü'], rotation=45, ha='right')

        # Layout'u sıkılaştır
        plt.tight_layout()

        # Grafiği tkinter'e ekle
        canvas = FigureCanvasTkAgg(fig, self.grafik_frame)
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        canvas.draw()

    def export_to_excel(self):
        if self.mismatched_data:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                try:
                    with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                        self.mismatched_data["gider_analizi"].to_excel(writer, sheet_name="Gider Analizi", index=False)
                        self.mismatched_data["gelir_analizi"].to_excel(writer, sheet_name="Gelir Analizi", index=False)
                    messagebox.showinfo("Başarılı", f"Rapor başarıyla kaydedildi: {save_path}")
                except Exception as e:
                    messagebox.showerror("Hata", f"Excel dosyası kaydedilirken hata oluştu:\n{str(e)}")