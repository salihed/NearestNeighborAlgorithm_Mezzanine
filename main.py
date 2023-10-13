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


def toplama(siparis_df, mezanin_df):
    toplama_listesi = []
    sira = 1

    # Katları sıralı bir şekilde dolaşalım.
    for z in range(1, 6):
        kat_mezanin_df = mezanin_df[mezanin_df['Depo adresi'].apply(lambda x: get_kat_as_number(x) == z)]
        koordinatlar = np.array([adres_to_koordinat(adres) for adres in kat_mezanin_df['Depo adresi']])
        nbrs = NearestNeighbors(n_neighbors=1).fit(koordinatlar)

        for index, row in siparis_df.iterrows():
            siparis_miktar = row['SİPARİŞ']
            urun = row['Malzeme']

            mezanin_urun_df = kat_mezanin_df[kat_mezanin_df['Ürün'] == urun]

            for urun_index, urun_row in mezanin_urun_df.iterrows():
                Depo_adresi = urun_row['Depo adresi']
                koordinat = adres_to_koordinat(Depo_adresi)
                if not koordinat:
                    continue
                _, indices = nbrs.kneighbors([koordinat])
                closest_index = indices[0][0]

                closest_koordinat = koordinatlar[closest_index]
                closest_adres_row = kat_mezanin_df[(kat_mezanin_df['Depo adresi'] == Depo_adresi)]

                miktar = min(siparis_miktar, closest_adres_row.iloc[0]['Miktar'])
                siparis_miktar -= miktar

                toplama_listesi.append({
                    'Toplama Sırası': sira,
                    'DIN NO': row['DIN NO'],
                    'Depo Adresi': Depo_adresi,
                    'Taşıma Birimi': urun_row['Taşıma birimi'],
                    'Malzeme': urun,
                    'Parti': urun_row['Parti'],
                    'Alınacak Miktar': miktar,
                    'Kat': get_kat_name(z),
                    'Koordinat': koordinat
                })
                sira += 1

                if siparis_miktar <= 0:
                    break

    return pd.DataFrame(toplama_listesi)


siparis_df = pd.read_excel("C:/Users/salih.dogrubak/Desktop/STD_YK_SIP.xlsx")
mezanin_df = pd.read_excel("C:/Users/salih.dogrubak/Desktop/Mezanin_STD.xlsx")
sonuc_df = toplama(siparis_df, mezanin_df)
sonuc_df.to_excel("C:/Users/salih.dogrubak/Desktop/ToplamaListesi_v3.xlsx", index=False)
