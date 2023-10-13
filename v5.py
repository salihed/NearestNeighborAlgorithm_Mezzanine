import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np

# Fonksiyonlar
def adres_to_koordinat(adres):
    koridor_dict = {'MZA': 1, 'MZB': 2, 'MZC': 3, 'MZD': 4, 'MZE': 5, 'MZF': 6, 'MZG': 7}

    if "." not in adres:
        return None

    koridor, x, y = adres.split('.')
    if koridor not in koridor_dict:
        return None

    z = get_kat_as_number(adres)
    return (z, koridor_dict[koridor], int(x), int(y))

def get_kat_as_number(Depo_adresi):
    y_ekseni = int(Depo_adresi.split('.')[2])
    if 1 <= y_ekseni <= 4:
        return 1
    elif 5 <= y_ekseni <= 8:
        return 2
    elif 9 <= y_ekseni <= 12:
        return 3
    elif 13 <= y_ekseni <= 16:
        return 4
    elif 17 <= y_ekseni <= 19:
        return 5
    else:
        return -1

def get_kat_name(z_number):
    kat_names = {
        1: "Zemin Kat",
        2: "1. Kat",
        3: "2. Kat",
        4: "3. Kat",
        5: "4. Kat"
    }
    return kat_names.get(z_number, "Bilinmeyen Kat")

def toplama(siparis_df, mezanin_df):
    toplama_listesi = []
    sira = 1
    toplanan_tasima_birimleri = set()

    koordinatlar = np.array([adres_to_koordinat(adres) for adres in mezanin_df['Depo adresi']])
    koordinatlar = sorted(koordinatlar, key=lambda x: (x[0], x[1], x[2], x[3]))

    for koordinat in koordinatlar:
        kat, koridor, x, y = koordinat
        Depo_adresi = f"MZ{chr(65 + koridor - 1)}.{x:03}.{y:02}"
        urunler_rows = mezanin_df[mezanin_df['Depo adresi'] == Depo_adresi]

        for _, urun_row in urunler_rows.iterrows():
            tasima_birimi = urun_row['Taşıma birimi']

            if tasima_birimi in toplanan_tasima_birimleri:
                continue

            urun = urun_row['Ürün']

            if urun not in siparis_df['Malzeme'].values:
                continue

            miktar = urun_row['Miktar']

            toplama_listesi.append({
                'Toplama Sırası': sira,
                'DIN NO': siparis_df[siparis_df['Malzeme'] == urun]['DIN NO'].iloc[0],
                'Depo Adresi': Depo_adresi,
                'Taşıma Birimi': tasima_birimi,
                'Malzeme': urun,
                'Parti': urun_row['Parti'],
                'Alınacak Miktar': miktar,
                'Kat': get_kat_name(kat),
                'Koordinat': koordinat
            })
            toplanan_tasima_birimleri.add(tasima_birimi)
            sira += 1

    return pd.DataFrame(toplama_listesi)


# Uygulama Sınıfı
class App:
    def __init__(self, master):
        self.master = master
        self.master.title("Mezanin En Yakın Adresten Toplama Uygulaması")

        self.siparis_dosya_yolu = None
        self.mezanin_dosya_yolu = None

        self.button1 = tk.Button(master, text="Sipariş Dosyası", command=self.load_siparis)
        self.button1.pack(pady=20)

        self.label1 = tk.Label(master, text="")
        self.label1.pack(pady=10)

        self.button2 = tk.Button(master, text="Mezanin Stokları", command=self.load_mezanin)
        self.button2.pack(pady=20)

        self.label2 = tk.Label(master, text="")
        self.label2.pack(pady=10)

        self.progress = ttk.Progressbar(master, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=20)

        self.run_button = tk.Button(master, text="Çalıştır", command=self.run_algorithm)
        self.run_button.pack(pady=20)

    def load_siparis(self):
        self.siparis_dosya_yolu = filedialog.askopenfilename(title="Sipariş Dosyası:")
        self.label1.config(text=self.siparis_dosya_yolu)

    def load_mezanin(self):
        self.mezanin_dosya_yolu = filedialog.askopenfilename(title="Mezanin Stokları:")
        self.label2.config(text=self.mezanin_dosya_yolu)

    def run_algorithm(self):
        if not self.siparis_dosya_yolu or not self.mezanin_dosya_yolu:
            messagebox.showerror("Hata", "Lütfen iki dosyayı da seçin.")
            return

        siparis_df = pd.read_excel(self.siparis_dosya_yolu)
        mezanin_df = pd.read_excel(self.mezanin_dosya_yolu)

        # Sürecin %50'sini tamamladığımızı varsayalım (Bunu daha sonra gerçek verilere göre ayarlayabilirsiniz)
        self.progress['value'] = 50
        self.master.update_idletasks()

        # Sonuc dosyasını oluşturma
        sonuc_df = toplama(siparis_df, mezanin_df)

        # Kaydedilecek dosya yolunu kullanıcıdan al
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path:  # Eğer kullanıcı iptal ederse veya bir konum seçmezse çıkış yap
            return
        sonuc_df.to_excel(save_path, index=False)

        self.progress['value'] = 100
        self.master.update_idletasks()

        messagebox.showinfo("Tamamlandı", "Toplama Dosyası oluşturuldu!")
        self.master.quit()

root = tk.Tk()
app = App(root)
root.mainloop()
