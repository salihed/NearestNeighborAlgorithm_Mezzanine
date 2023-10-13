import pandas as pd
import numpy as np
from sklearn.neighbors import NearestNeighbors


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

#DEĞİŞTİRİLEN
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


siparis_df = pd.read_excel("C:/Users/salih.dogrubak/Desktop/STD_YK_SIP.xlsx")
mezanin_df = pd.read_excel("C:/Users/salih.dogrubak/Desktop/Mezanin_STD.xlsx")
sonuc_df = toplama(siparis_df, mezanin_df)
sonuc_df.to_excel("C:/Users/salih.dogrubak/Desktop/ToplamaListesi_v1.xlsx", index=False)
