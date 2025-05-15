import tkinter as tk
from tkinter import ttk
from desi_uyusmazligi_tespit import DesiUyusmazligiTespit
from trendyol_cari_ekstre_raporlayici import TrendyolCariEkstreRaporlayici
from trendyol_siparis_analiz import TrendyolSiparisAnaliz


class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Trendyol Araçları")
        self.root.geometry("1200x800")

        # Tema ve Stil Ayarları
        self.setup_styles()

        # Sekme (Notebook) widget'ı oluştur
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Durum Çubuğu (status_bar) Tanımlama
        self.status_bar = ttk.Label(self.root, text="Hazır", relief=tk.SUNKEN, anchor="w", font=("Arial", 10))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Uygulamalar için sekmeler oluştur
        self.create_tab("🏷️ Desi Uyuşmazlığı Tespit", DesiUyusmazligiTespit)
        self.create_tab("📊 Trendyol Cari Ekstre Raporlayıcı", TrendyolCariEkstreRaporlayici)
        self.create_tab("📈 Trendyol Sipariş Analiz", TrendyolSiparisAnaliz)

    def setup_styles(self):
        """Tema ve stilleri ayarla"""
        style = ttk.Style(self.root)
        style.theme_use("default")

        # Sekme arka planı
        style.configure("TNotebook", background="#f0f0f0")
        style.configure("TNotebook.Tab", font=("Arial", 12, "bold"), padding=[10, 5])

        # Durum çubuğu stili
        style.configure("TLabel", background="#e0e0e0", font=("Arial", 10))

    def create_tab(self, title, app_class):
        """Yeni bir sekme oluştur ve uygulamayı yerleştir."""
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text=title)  # Sekmeyi Notebook'a ekle

        # Sekme içeriği
        try:
            app_class(tab_frame)  # Uygulamayı sekme içine başlat
            self.update_status(f"{title} sekmesi başarıyla yüklendi.")
        except Exception as e:
            self.update_status(f"{title} sekmesi yüklenirken bir hata oluştu: {str(e)}")

    def update_status(self, message):
        """Durum çubuğunu güncelle"""
        self.status_bar.config(text=message)


if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()