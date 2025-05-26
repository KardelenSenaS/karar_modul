#!/usr/bin/env python
# coding: utf-8

# In[32]:


import pandas as pd

# Kodlar aracılığıyla hesapladığımız, modulümüzde kullanacağımız rasyolardan oluşturduğumuz Excel'den oranları okumak için;
df = pd.read_excel("C:/Users/HP/Desktop/finansal_rasyolar.xlsx")
print(df)


# In[33]:


# Karar destek modülünü yazarken çağırabilmek için oranları sözlük yapısına çeviriyoruz;
oranlar = dict(zip(df["Oran"], df["Değer"]))
print(oranlar)


# In[34]:


# Örneğin:
oranlar["Cari Oran"]


# In[35]:


# Değerleri temiz sayıya çevirelim, yani string olan veriler matematiksel işlem yapılabilmesi için float edilsin;
def temizle_sayi(x):
    if isinstance(x, str):
        x = x.replace(" TL", "").replace("−", "-").replace(" ", "")
        if "," in x and "." not in x:
            # Ondalık formatı (örnek: "1,52") → "1.52"
            x = x.replace(",", ".")
        x = x.replace(",", "")  # Her ihtimale karşı kalan virgül varsa temizle
    return float(x)


# In[36]:


# Karar destek modülününde kullanacağımız 9 rasyoyu sayıya çevirmek için (hata payı bırakmadan);
cari_oran = temizle_sayi(oranlar["Cari Oran"])
asit_test_orani = temizle_sayi(oranlar["Asit Test Oranı"])
finansman_suresi = temizle_sayi(oranlar["Finansman Süresi"])
kaldirac_orani = temizle_sayi(oranlar["Kaldıraç Oranı"])
ozkaynak_orani = temizle_sayi(oranlar["Öz Kaynak Oranı"])
aktif_karlilik = temizle_sayi(oranlar["Aktif Karlılık"])
ekonomik_rantabilite = temizle_sayi(oranlar["Ekonomik Rantabilite"])
ebitda_marji = temizle_sayi(oranlar["EBITDA Marjı"])
net_borc_ebitda = temizle_sayi(oranlar["Net Finansal Borç / EBITDA"])


# In[37]:


# Çarpan etkisi ile rasyoların ağırlıklı puan ve nihai hesaplanacak nota katkısını belirleyelim; 

# 1. Cari Oran Ağırlıklı Puan Hesabı
def puanla_cari_oran(deger):
    
    if deger < 0.50: #ciddilikiditesorunu
        return 0
    elif deger < 0.80: #kritik
        return 3
    elif deger < 1.20: #idareder
        return 6
    elif deger < 1.50: #makul
        return 8
    elif deger <= 2.70: #saglıklı
        return 10
    else:
        return 7  #aşırılikidite  -  kaynaklar verimsiz kullanılıyor olabilir

    
    # Cari Oranın %12 Çarpan ile Not'a Katkısı
cari_oran_puani = puanla_cari_oran(cari_oran)
cari_oran_katkisi = cari_oran_puani * 0.12

print(f"Cari Oran: {cari_oran} → Puan: {cari_oran_puani}, Katkı: {cari_oran_katkisi}")


# In[38]:


# 2. Asit Test Oranı Ağırlıklı Puan Hesabı
def puanla_asit_test_orani(deger):
    if deger < 0.30: #ciddilikiditesorunu
        return 0
    elif deger < 0.60: #kritik
        return 3
    elif deger < 0.85:  #idareder
        return 5
    elif deger < 1.20:  #makul
        return 8
    elif deger <= 2.20: #saglıklı
        return 10
    else:
        return 6  #aşırılikidite -  kaynaklar verimsiz kullanılıyor olabilir


   # Asit Test Oranın %10 Çarpan ile Not'a Katkısı
asit_test_orani_puani = puanla_asit_test_orani(asit_test_orani)
asit_test_orani_katkisi = asit_test_orani_puani * 0.10

print(f"Asit Test Oranı: {asit_test_orani} → Puan: {asit_test_orani_puani}, Katkı: {asit_test_orani_katkisi }")


# In[39]:


# 3. Finansman Süresi Ağırlıklı Puan Hesabı
def puanla_finansman_suresi(gun):
    if gun <= 0:  #mükemmel - tedarikçiye borç ödeme süresi alacak tahsilinden daha uzun
        return 10
    elif gun < 30:  #güclünakitdöngüsü
        return 9
    elif gun < 60:  #makulfinansmansüresi - tahsil süresine göre 1-2 ay arası finansmana ihtiyaç duyuyor
        return 8
    elif gun < 90:  #artandıskaynakihtiyacı
        return 5
    elif gun < 120:  #zayıfnakitakıs
        return 3
    else:
        return 0  #kritik – 4 aydan uzun süre dış kaynak ihityac


   # Finansman Süresi %7 Çarpan ile Not'a Katkısı
finansman_suresi_puani = puanla_finansman_suresi(finansman_suresi)
finansman_suresi_katkisi = finansman_suresi_puani * 0.07

print(f"Finansman Süresi: {finansman_suresi} gün → Puan: {finansman_suresi_puani}, Katkı: {finansman_suresi_katkisi}")


# In[40]:


# 4. Kaldıraç Oranı Ağırlıklı Puan Hesbı
def puanla_kaldirac_orani(deger):
    if deger < 0.30:  #cokdüsükborc - kullanılan kaynaklar çoğunlukla özkaynak
        return 10
    elif deger < 0.50:  #güclüsermaye
        return 9
    elif deger < 0.70:  #makulborc
        return 7
    elif deger < 0.90:  #artanborcluluk
        return 5
    elif deger < 1.00:  #yüksekborc
        return 3
    else:
        return 0  #kritik – borç özkaynaktan fazla


  # Kaldıraç Oranı %17 Çarpan ile Not'a Katkısı
kaldirac_orani_puani = puanla_kaldirac_orani(kaldirac_orani)
kaldirac_orani_katkisi = kaldirac_orani_puani * 0.17

print(f"Kaldıraç Oranı: {kaldirac_orani} → Puan: {kaldirac_orani_puani}, Katkı: {kaldirac_orani_katkisi }")


# In[41]:


# 5. Özkaynak Oranı Ağırlıklı Puan Hesabı
def puanla_ozkaynak_orani(deger):
    if deger >= 0.60:  #cokgüclüözkaynak
        return 10
    elif deger >= 0.40:  #saglıklı
        return 9
    elif deger >= 0.20:  #makul
        return 6
    elif deger >= 0.10:  #zayıfsermaye
        return 4
    elif deger >= 0.5:  #kritigeyakın
        return 2
    else:
        return 0  #kritik – aktifler tamamen yabancı kaynaklar ile finanse ediliyor


  # Özkaynak Oranı %8 Çarpan ile Not'a Katkısı
ozkaynak_orani_puani = puanla_ozkaynak_orani(ozkaynak_orani)
ozkaynak_orani_katkisi = ozkaynak_orani_puani * 0.08  # %8 ağırlık

print(f"Öz Kaynak Oranı: {ozkaynak_orani} → Puan: {ozkaynak_orani_puani}, Katkı: {ozkaynak_orani_katkisi}")


# In[42]:


# 6. Aktif Karlılık Ağırlıklı Puan Hesabı
def puanla_aktif_karlilik(deger):
    if deger >= 15:  #mükemmel
        return 10
    elif deger >= 10:  #güclü
        return 9
    elif deger >= 6:  #orta
        return 8
    elif deger >= 3:  #zayıf
        return 6
    elif deger >= 0:  #sınırdakarlılık
        return 4
    else:
        return 0  #kritik - sahip olunan aktifler ile kar üretilemiyor


    # Aktif Karlılık %10 Çarpan ile Not'a Katkısı
aktif_karlilik_puani = puanla_aktif_karlilik(aktif_karlilik)
aktif_karlilik_katkisi = aktif_karlilik_puani * 0.10

print(f"Aktif Karlılık: {aktif_karlilik}% → Puan: {aktif_karlilik_puani}, Katkı: {aktif_karlilik_katkisi}")


# In[43]:


# 7. Ekonomik Rantabilite Ağırlıklı Puan Hesabı
def puanla_ekonomik_rantabilite(deger):
    if deger >= 20:  #üstdüzey
        return 10
    elif deger >= 10:  #yüksekperformans
        return 8
    elif deger >= 5:  #ortalama
        return 6
    elif deger >= 2:  #zayıf
        return 4
    elif deger >= 0:  #cokdüsükyatırımdönüsü
        return 2
    else:
        return 0  #negatif - yatırımlar değer üretmiyor


    # Ekonomik Rantabilite %6 Çarpan ile Not'a Katkısı
ekonomik_rantabilite_puani = puanla_ekonomik_rantabilite(ekonomik_rantabilite)
ekonomik_rantabilite_katkisi = ekonomik_rantabilite_puani * 0.06

print(f"Ekonomik Rantabilite: {ekonomik_rantabilite}% → Puan: {ekonomik_rantabilite_puani}, Katkı: {ekonomik_rantabilite_katkisi}")


# In[44]:


# 8. EBITDA Marjı Ağırlıklı Puan Hesabı
def puanla_ebitda_marji(deger):
    if deger >= 20:  #mükemmelverimlilik
        return 10 
    elif deger >= 15:  #güclükarlılık
        return 9
    elif deger >= 10:  #makul
        return 8
    elif deger >= 5:  #düsükverim
        return 6
    elif deger >= 0:  #kritik
        return 2
    else:
        return 0  #faaliyetlerdenzarar


  # EBITDA Marjı %13 Çarpan ile Not'a Katkısı
ebitda_marji_puani = puanla_ebitda_marji(ebitda_marji)
ebitda_marji_katkisi = ebitda_marji_puani * 0.13

print(f"EBITDA Marjı: {ebitda_marji}% → Puan: {ebitda_marji_puani}, Katkı: {ebitda_marji_katkisi}")


# In[45]:


# 9. Net Finansal Borç / EBITDA Ağırlıklı Puan Hesabı
def puanla_net_borc_ebitda(deger):
    if deger < 1.0:  #mükemmel - borç çok hızlı ödeniyor
        return 10
    elif deger < 2.0:  #güclü
        return 9
    elif deger < 3.0:  #makul
        return 7
    elif deger < 4.0:  #artanrisk
        return 5
    elif deger < 6.0:  #yüksekrisk
        return 2
    else:
        return 0  # kritik – elde edilen kar ile borç ödemesi 6 yıldan uzun sürece


   # Net Finansal Borç / EBITDA %17 Çarpan ile Not'a Katkısı
net_borc_ebitda_puani = puanla_net_borc_ebitda(net_borc_ebitda)
net_borc_ebitda_katkisi = net_borc_ebitda_puani * 0.17

print(f"Net Borç / EBITDA: {net_borc_ebitda} → Puan: {net_borc_ebitda_puani}, Katkı: {net_borc_ebitda_katkisi}")


# In[46]:


# Parameterik olan çarpanlar tanımlanır, toplamı 1.00 olmalıdır;
agirliklar = {
    "cari_oran": 0.12,
    "asit_test": 0.10,
    "finansman_suresi": 0.07,
    "kaldirac": 0.17,
    "ozkaynak_orani": 0.08,
    "aktif_karlilik": 0.10,
    "ekonomik_rantabilite": 0.06,
    "ebitda_marji": 0.13,
    "net_borc_ebitda": 0.17
}

toplam_agirlik = sum(agirliklar.values())
print(f"Ağırlıkların Toplamı: {toplam_agirlik:.4f}")

if abs(toplam_agirlik - 1.00) > 0.001:
    print("UYARI: Ağırlıkların toplamı %100 değil! FiNo sonucu hatalı olabilir.")


# In[47]:


# Hesaplanan rasyolar ve çarpanlarından elde eden katsayı puanlarının toplamından 0-10 arası olacak FiNo'yu otomatik eldesi;
def hesapla_fino_otomatik():
    toplam = 0
    for isim, deger in globals().items(): 
        if isim.endswith("_katkisi") and isinstance(deger, (int, float)):
            toplam += deger
    return round(toplam, 2)

fino = hesapla_fino_otomatik()
print("FiNo:", fino)


# In[48]:


# FiNo’ya göre limit talebinin oransal karşılığını belirlemek için;

def kredi_katsayisi(fino):
    if fino >= 9.0:
        return 1.00
    elif fino >= 8.0:
        return 0.80
    elif fino >= 7.0:
        return 0.70
    elif fino >= 5.0:
        return 0.50
    elif fino >= 4.0:
        return 0.40
    else:
        return 0.00


# In[49]:


# Okutacağımız verinin dosya yolunu tanımladık (Finansal verileri kullanabilmek için);
dosya_yolu = r"C:\Users\HP\Desktop\KAP VERİLERİ\ASELS_1395801_2024_4.xls"

# KAP'tan temin edilen .xls uzantılı ancak içeriği HTML olan dosyayı okutmak ve liste halinde dfs değişkenine atanması için;
dfs = pd.read_html(dosya_yolu)

# KAP' tan temin edilen dosya içerisinde bilanço ve gelir tablosu alanlarının otomatik bulunabilmesi için; 

# Başlangıçta None olarak tanımlıyoruz
df_bilanco = None
df_gelir = None

# Tüm tabloları sırayla gez;
for i, df in enumerate(dfs):
    try:
        # Tablonun ilk birkaç satırını düz metne çevir ve küçük harfe indir
        metin = " ".join(df.astype(str).head(5).values.flatten()).lower()
        
        # Bilanço tespiti
        if "bilanço" in metin and df_bilanco is None:
            df_bilanco = df
            bilanco_index = i
        
        # Gelir tablosu tespiti
        if "gelir tablosu" in metin and df_gelir is None:
            df_gelir = df
            gelir_index = i

        # İkisini de bulduysa döngüyü durdur, tekrar eden isim olduğu durumunda ilkini seçmek için;
        if df_bilanco is not None and df_gelir is not None:
            break

    except:
        continue
# Check        
print(f"✅ Bilanço tablosu: dfs[{bilanco_index}]")
print(f"✅ Gelir tablosu: dfs[{gelir_index}]")

# Bilanço ve gelir tablolarını bulduk, şimdi tablolar içerisindeki değerlere 
#kalem isimleri ile erişebilmek için tabloları "veriler" sözlüğüne dönüştüreceğiz ;

# 1. Bilanço ve gelir tablosundan doğru sütunları seç (Hesap Adı, Değer)
bilanco_clean = df_bilanco[[1, 3]].dropna().reset_index(drop=True)
gelir_clean = df_gelir[[1, 3]].dropna().reset_index(drop=True)

# 2. Her tabloyu sözlük yapısına çevir
sozluk_bilanco = dict(zip(bilanco_clean[1], bilanco_clean[3]))
sozluk_gelir = dict(zip(gelir_clean[1], gelir_clean[3]))

# 3. İki sözlüğü birleştir
veriler = {**sozluk_bilanco, **sozluk_gelir}

# 4. Sayıları Türkçe formatından temizle (1.234.567,89 → 1234567.89)
def temizle_sayi(x):
    if isinstance(x, str):
        return float(x.replace(".", "").replace(",", ".").replace("−", "-").replace(" ", ""))
    return x

for k in veriler:
    try:
        veriler[k] = temizle_sayi(veriler[k])
    except:
        continue
        
# 5. Sözlük örneği (ilk 10 anahtarı ve değeri göster)
veriler_ornek = {k: veriler[k] for k in list(veriler)[:10]}
veriler_ornek


# In[50]:


veriler["Net Dönem Karı veya Zararı"]


# In[51]:


veriler["Hasılat"]


# In[52]:


veriler["TOPLAM ÖZKAYNAKLAR"]


# In[53]:


veriler["TOPLAM VARLIKLAR"]


# In[54]:


mevcut_kredi_borcu = veriler["Kısa Vadeli Borçlanmalar"] + veriler['Uzun Vadeli Borçlanmaların Kısa Vadeli Kısımları'] + veriler["Uzun Vadeli Borçlanmalar"]
print(mevcut_kredi_borcu)


# In[55]:


net_finansal_borc = veriler["Kısa Vadeli Borçlanmalar"] + veriler['Uzun Vadeli Borçlanmaların Kısa Vadeli Kısımları'] + veriler["Uzun Vadeli Borçlanmalar"] - veriler["Nakit ve Nakit Benzerleri"]
print(net_finansal_borc)


# In[56]:


oranlar["EBITDA"]


# In[58]:


def formatla_tl(deger):
    return f"{deger:,.0f}".replace(",", ".") + " TL"


# In[59]:


#Finansal veriler okuttuğumuz Excel'de.000 basamak eksik yer almaktadır (Bu bilgi Bağımsız Denetim Raporu'nda yer almaktadır)

#Verilerin orjinal haline erişebilmek için;
veriler = {
    "Hasılat": 120205594.0,
    "TOPLAM ÖZKAYNAKLAR": 141359149.0,
    "TOPLAM VARLIKLAR" : 242797511.0,
    "net_finansal_borc": 15925058.0,
    "mevcut_kredi_borcu" : 32562322.0,
    "EBITDA": 22328470.0
}

#x1000 yapılmıştır
for kalem, deger in veriler.items():
    print(f"{kalem}: {formatla_tl(deger*1000)}")


# In[66]:


def hesapla_teorik_limit(veriler: dict):
    ebitda   = veriler["EBITDA"] * 1000
    net_borc = veriler["net_finansal_borc"] * 1000
    aktif    = veriler["TOPLAM VARLIKLAR"] * 1000
    hasilat  = veriler["Hasılat"] * 1000
    ozkaynak = veriler["TOPLAM ÖZKAYNAKLAR"] * 1000

    limit = ((ebitda * 4) - net_borc) + (aktif * 0.4) + (hasilat * 0.2) + (ozkaynak * 1.5)
    return round(limit, 2)

# Formatlama
def formatla_tl(sayi):
    return f"{sayi:,.2f} TL".replace(",", ".")

# Kullanım:
teorik_limit = hesapla_teorik_limit(veriler)
print("Teorik Limit:", formatla_tl(teorik_limit))


# In[67]:


def hesapla_verilebilecek_limit(veriler: dict):
    fino = hesapla_fino_otomatik()
    teorik_limit = hesapla_teorik_limit(veriler)
    katsayi = kredi_katsayisi(fino)
    return round(teorik_limit * katsayi, 2)

limit = hesapla_verilebilecek_limit(veriler)
print("Verilebilecek Limit:", formatla_tl(limit))


# In[68]:


def karar_modulu_yurut(talep_edilen_yeni_kredi):
    try:
        fino = hesapla_fino_otomatik()
        teorik_limit = hesapla_teorik_limit(veriler)
        katsayi = kredi_katsayisi(fino)
        verilebilecek_limit = teorik_limit * katsayi
        mevcut_kredi_borcu = veriler["mevcut_kredi_borcu"] * 1000
        toplam_kredi_yuku = mevcut_kredi_borcu + talep_edilen_yeni_kredi

        return (
            f"FiNo Skoru: {fino:.2f} / 10.00\n"
            f"Teorik Limit: {verilebilecek_limit:,.0f} TL (katsayılı)\n"
            f"Mevcut Kredi Borcu: {mevcut_kredi_borcu:,.0f} TL\n"
            f"Yeni Talep: {talep_edilen_yeni_kredi:,.0f} TL\n"
            f"Toplam Kredi Yükü: {toplam_kredi_yuku:,.0f} TL\n" +
            ("RED: Firma notu çok düşük, kredi verilmemeli." if katsayi == 0 else
             "ONAY: Toplam risk kabul edilebilir düzeyde." if toplam_kredi_yuku <= verilebilecek_limit else
             "REVİZE: Toplam risk sınırı aşıldı. Talep revize edilmeli veya teminat istenmeli.")
        )

    except Exception as e:
        import traceback
        return f"HATA: {str(e)}\n{traceback.format_exc()}"


# In[ ]:




