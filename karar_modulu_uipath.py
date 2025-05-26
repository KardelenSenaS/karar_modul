
def karar_modulu_yurut(talep_edilen_yeni_kredi):
    try:
        import pandas as pd
        import traceback

        # Oran dosyasını oku
        df = pd.read_excel("C:/Users/HP/Desktop/finansal_rasyolar.xlsx")
        oranlar = dict(zip(df["Oran"], df["Değer"]))

        def temizle_sayi(x):
            if isinstance(x, str):
                x = x.replace(" TL", "").replace("−", "-").replace(" ", "")
                if "," in x and "." not in x:
                    x = x.replace(",", ".")
                x = x.replace(",", "")
            return float(x)

        # 9 temel oranı oku ve temizle
        cari_oran = temizle_sayi(oranlar["Cari Oran"])
        asit_test_orani = temizle_sayi(oranlar["Asit Test Oranı"])
        finansman_suresi = temizle_sayi(oranlar["Finansman Süresi"])
        kaldirac_orani = temizle_sayi(oranlar["Kaldıraç Oranı"])
        ozkaynak_orani = temizle_sayi(oranlar["Öz Kaynak Oranı"])
        aktif_karlilik = temizle_sayi(oranlar["Aktif Karlılık"])
        ekonomik_rantabilite = temizle_sayi(oranlar["Ekonomik Rantabilite"])
        ebitda_marji = temizle_sayi(oranlar["EBITDA Marjı"])
        net_borc_ebitda = temizle_sayi(oranlar["Net Finansal Borç / EBITDA"])

        # Her oran için puan fonksiyonları
        def puanla_cari_oran(x): return 0 if x < 0.50 else 3 if x < 0.80 else 6 if x < 1.20 else 8 if x < 1.50 else 10 if x <= 2.70 else 7
        def puanla_asit_test_orani(x): return 0 if x < 0.30 else 3 if x < 0.60 else 5 if x < 0.85 else 8 if x < 1.20 else 10 if x <= 2.20 else 6
        def puanla_finansman_suresi(x): return 10 if x <= 0 else 9 if x < 30 else 8 if x < 60 else 5 if x < 90 else 3 if x < 120 else 0
        def puanla_kaldirac_orani(x): return 10 if x < 0.30 else 9 if x < 0.50 else 7 if x < 0.70 else 5 if x < 0.90 else 3 if x < 1.00 else 0
        def puanla_ozkaynak_orani(x): return 10 if x >= 0.60 else 9 if x >= 0.40 else 6 if x >= 0.20 else 4 if x >= 0.10 else 2 if x >= 0.05 else 0
        def puanla_aktif_karlilik(x): return 10 if x >= 15 else 9 if x >= 10 else 8 if x >= 6 else 6 if x >= 3 else 4 if x >= 0 else 0
        def puanla_ekonomik_rantabilite(x): return 10 if x >= 20 else 8 if x >= 10 else 6 if x >= 5 else 4 if x >= 2 else 2 if x >= 0 else 0
        def puanla_ebitda_marji(x): return 10 if x >= 20 else 9 if x >= 15 else 8 if x >= 10 else 6 if x >= 5 else 2 if x >= 0 else 0
        def puanla_net_borc_ebitda(x): return 10 if x < 1.0 else 9 if x < 2.0 else 7 if x < 3.0 else 5 if x < 4.0 else 2 if x < 6.0 else 0

        # Ağırlıklı katkılar
        katkilar = {
            "cari": puanla_cari_oran(cari_oran) * 0.12,
            "asit": puanla_asit_test_orani(asit_test_orani) * 0.10,
            "finansman": puanla_finansman_suresi(finansman_suresi) * 0.07,
            "kaldirac": puanla_kaldirac_orani(kaldirac_orani) * 0.17,
            "ozkaynak": puanla_ozkaynak_orani(ozkaynak_orani) * 0.08,
            "aktif": puanla_aktif_karlilik(aktif_karlilik) * 0.10,
            "rant": puanla_ekonomik_rantabilite(ekonomik_rantabilite) * 0.06,
            "ebitda_marji": puanla_ebitda_marji(ebitda_marji) * 0.13,
            "net_borc": puanla_net_borc_ebitda(net_borc_ebitda) * 0.17
        }

        fino = round(sum(katkilar.values()), 2)

        def kredi_katsayisi(fino):
            return 1.00 if fino >= 9 else 0.80 if fino >= 8 else 0.70 if fino >= 7 else 0.50 if fino >= 5 else 0.40 if fino >= 4 else 0.00

        veriler = {
            "Hasılat": 120205594.0,
            "TOPLAM ÖZKAYNAKLAR": 141359149.0,
            "TOPLAM VARLIKLAR": 242797511.0,
            "net_finansal_borc": 15925058.0,
            "mevcut_kredi_borcu": 32562322.0,
            "EBITDA": 22328470.0
        }

        def hesapla_teorik_limit(veriler: dict):
            ebitda   = veriler["EBITDA"] * 1000
            net_borc = veriler["net_finansal_borc"] * 1000
            aktif    = veriler["TOPLAM VARLIKLAR"] * 1000
            hasilat  = veriler["Hasılat"] * 1000
            ozkaynak = veriler["TOPLAM ÖZKAYNAKLAR"] * 1000
            limit = ((ebitda * 4) - net_borc) + (aktif * 0.4) + (hasilat * 0.2) + (ozkaynak * 1.5)
            return round(limit, 2)

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
        return f"HATA: {str(e)}\n{traceback.format_exc()}"
