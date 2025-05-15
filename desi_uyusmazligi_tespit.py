import pandas as pd
from tkinter import Button, Label, filedialog, messagebox
from tkinter.ttk import Treeview


class DesiUyusmazligiTespit:
    def __init__(self, root):
        self.root = root
        self.delivered_file = None
        self.returned_file = None
        self.mismatched_data = None  # Uyuşmazlıkları saklamak için

        # Teslim edilen siparişler butonu
        self.delivered_button = Button(root, text="Teslim Edilen Siparişler Excellini Yükle", command=self.load_delivered_excel)
        self.delivered_button.pack(pady=10)

        # İade edilen siparişler butonu
        self.returned_button = Button(root, text="İade Edilen Siparişler Excellini Yükle", command=self.load_returned_excel)
        self.returned_button.pack(pady=10)

        # Sonuç gösterimi
        self.result_label = Label(root, text="Sonuçlar:")
        self.result_label.pack(pady=10)

        self.result_tree = Treeview(root, columns=("Sipariş Numarası", "Teslim Desi", "İade Desi"), show="headings")
        self.result_tree.heading("Sipariş Numarası", text="Sipariş Numarası")
        self.result_tree.heading("Teslim Desi", text="Teslim Desi")
        self.result_tree.heading("İade Desi", text="İade Desi")
        self.result_tree.pack(pady=10)

        # Özet gösterimi
        self.summary_label = Label(root, text="", fg="blue")
        self.summary_label.pack(pady=10)

        # Karşılaştırma ve Excel İndir butonları
        self.compare_button = Button(root, text="Karşılaştır ve Sonuçları Göster", command=self.compare_excels)
        self.compare_button.pack(pady=10)

        self.export_button = Button(root, text="Raporu Excel Olarak İndir", command=self.export_to_excel, state="disabled")
        self.export_button.pack(pady=10)

    def load_delivered_excel(self):
        self.delivered_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.delivered_file:
            messagebox.showinfo("Yükleme Başarılı", "Teslim edilen siparişler excellini yüklediniz.")

    def load_returned_excel(self):
        self.returned_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.returned_file:
            messagebox.showinfo("Yükleme Başarılı", "İade edilen siparişler excellini yüklediniz.")

    def compare_excels(self):
        if not self.delivered_file or not self.returned_file:
            messagebox.showwarning("Eksik Dosya", "Lütfen her iki Excel dosyasını da yükleyin!")
            return

        try:
            # Teslim edilen siparişler Excel'ini yükle
            delivered_df = pd.read_excel(self.delivered_file)
            delivered_data = delivered_df[['Sipariş Numarası', 'Kargodan alınan desi']]

            # İade edilen siparişler Excel'ini yükle
            returned_df = pd.read_excel(self.returned_file)
            returned_data = returned_df[['Sipariş Numarası', 'İade Desi']]

            # Desi bilgisi olmayan siparişleri filtrele
            delivered_data = delivered_data[delivered_data['Kargodan alınan desi'].notna()]
            returned_data = returned_data[returned_data['İade Desi'].notna()]

            # Uyuşmayan siparişleri bul
            merged_data = pd.merge(
                delivered_data, 
                returned_data, 
                on='Sipariş Numarası', 
                suffixes=('_Teslim', '_İade')
            )
            mismatched = merged_data[merged_data['Kargodan alınan desi'] != merged_data['İade Desi']]

            # Sonuçları tabloya ekle
            self.result_tree.delete(*self.result_tree.get_children())  # Mevcut veriyi temizle
            if mismatched.empty:
                messagebox.showinfo("Sonuç", "Herhangi bir uyuşmazlık bulunamadı.")
                self.export_button.config(state="disabled")
                self.summary_label.config(text="")  # Özet bilgisini temizle
            else:
                self.mismatched_data = mismatched  # Uyuşmazlıkları kaydet
                for index, row in mismatched.iterrows():
                    self.result_tree.insert("", "end", values=(row['Sipariş Numarası'], row['Kargodan alınan desi'], row['İade Desi']))

                # Özet bilgisi hesaplama
                mismatched['Desi Farkı'] = abs(mismatched['Kargodan alınan desi'] - mismatched['İade Desi'])
                total_mismatches = len(mismatched)
                avg_difference = mismatched['Desi Farkı'].mean()
                max_difference = mismatched['Desi Farkı'].max()
                min_difference = mismatched['Desi Farkı'].min()

                # Özet bilgisini göster
                summary_text = (
                    f"Toplam Uyuşmazlık: {total_mismatches}\n"
                    f"Ortalama Desi Farkı: {avg_difference:.2f}\n"
                    f"En Büyük Desi Farkı: {max_difference:.2f}\n"
                    f"En Küçük Desi Farkı: {min_difference:.2f}"
                )
                self.summary_label.config(text=summary_text)
                self.export_button.config(state="normal")

        except Exception as e:
            messagebox.showerror("Hata", f"Bir hata oluştu: {e}")

    def export_to_excel(self):
        if self.mismatched_data is not None:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                try:
                    self.mismatched_data.to_excel(save_path, index=False)
                    messagebox.showinfo("Başarılı", f"Rapor başarıyla kaydedildi: {save_path}")
                except Exception as e:
                    messagebox.showerror("Hata", f"Excel dosyası kaydedilirken bir hata oluştu: {e}")