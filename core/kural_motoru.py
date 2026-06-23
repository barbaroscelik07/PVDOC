"""
Kural motoru v3 — Tablo 8'den Tablo 6/7/9 türetir.

tablet_ipk sözlüğü kullanıcıdan gelen ara-aşama spesifikasyonları:
  "Görünüş_Karışım", "Görünüş_Tablet", "Ortalama Ağırlık_Tablet",
  "Kalınlık", "Çap", "Sertlik"
Aşınma sabit "Maksimum %1.0"; Tablet Dağılma sabit "Maksimum 15 dakika".
"""

from __future__ import annotations

import copy

from core.models import Test, Spesifikasyon, LimitTuru, TabloTipi

OP_NO = {"Karıştırma": 2, "Tablet Baskı": 3, "Film Kaplama": 4,
         "Dolum": 3, "Blisterleme": 5}

SABIT_KARISIM_SPEK = "%85 – %115"
BILGI = "Bilgi amaçlıdır."
ASINMA_SPEK = "Maksimum %1.0"
TABLET_DAGILMA = "Maksimum 15 dakika"
SIZDIRMAZLIK_SPEK = "Sızdırmamalıdır."


def _norm(s: str) -> str:
    tr = str.maketrans({"ı": "i", "İ": "i", "ş": "s", "Ş": "s", "ğ": "g",
                        "ç": "c", "ö": "o", "ü": "u", "I": "i"})
    return (s or "").translate(tr).lower().strip()


def _etken_adlari(etkenler, testler):
    adlar = [em.ad for em in etkenler]
    if not adlar:
        bulunan = []
        for t in testler:
            if _norm(t.ad).startswith("etkin madde"):
                parca = t.ad.split()
                if len(parca) >= 3:
                    ad = " ".join(parca[:3])
                    if ad not in bulunan:
                        bulunan.append(ad)
        adlar = bulunan
    return adlar or ["Etkin madde 1"]


def _bul(testler, *anahtarlar, etken=None):
    for t in testler:
        n = _norm(t.ad)
        if all(_norm(a) in n for a in anahtarlar):
            if etken is None or _norm(etken) in n:
                return t
    return None


def _yeni(ad, op, tip, spek_metni="", *, ipk=False, yildiz=False,
          mikro=False, alt_satirlar=None, kaynak_spek=None,
          ac1="", as1="", ac2="", as2=""):
    if kaynak_spek is not None:
        sp = copy.deepcopy(kaynak_spek)
    else:
        sp = Spesifikasyon(limit_turu=LimitTuru.ARALIK, spesifikasyon_metni=spek_metni)
    return Test(ad=ad, operasyon=op, operasyon_no=OP_NO.get(op, 0),
                tablo_tipi=tip, ipk=ipk, yildizli=yildiz, mikrobiyolojik=mikro,
                spesifikasyon=sp, alt_satirlar=list(alt_satirlar or []),
                aciklama_etiketi=ac1, aciklama_spek=as1,
                aciklama2_etiketi=ac2, aciklama2_spek=as2)


def test_taninir_mi(ad: str) -> bool:
    """Bu test adı film tablet kural haritasında tanınıyor mu?"""
    n = _norm(ad)
    anahtarlar = ["gorunus", "elek", "karisim tekduzeligi", "ortalama agirlik",
                  "agirlik tekduzeligi", "agirlik sapmasi", "sertlik", "kalinlik",
                  "cap", "asinma", "dagilma", "teshis", "miktar tayini",
                  "dissol", "ilgili bilesik", "mikrobiyolojik", "sizdirmazlik"]
    return any(a in n for a in anahtarlar)


def taninmayan_testler(bitmis_testler) -> list:
    """Kuralda tanınmayan testlerin adlarını döndürür (etken öneki çıkarılmış)."""
    out = []
    for t in bitmis_testler:
        if not test_taninir_mi(t.ad):
            if t.ad not in out:
                out.append(t.ad)
    return out


def _ozel_kuralli_ekle(cikti, bitmis_testler, ozel_kurallar):
    """Özel kuralı olan tanımsız testleri ilgili aşamalara dağıtır."""
    if not ozel_kurallar:
        return
    for t in bitmis_testler:
        kural = ozel_kurallar.get(t.ad)
        if not kural:
            continue
        spek = kural.get("spek") or (t.spesifikasyon.spesifikasyon_metni
                                     or t.spesifikasyon.metni_olustur())
        for op in kural.get("asamalar", []):
            yildiz = op in kural.get("yildiz", [])
            cikti.append(_yeni(t.ad, op, TabloTipi.TEK_SONUC, spek,
                               ipk=kural.get("ipk", False), yildiz=yildiz))


def turet(bitmis_testler, etkin_maddeler, operasyonlar,
          cift_katman=False, tablet_ipk=None, ozel_test_kurallari=None):
    tablet_ipk = tablet_ipk or {}
    etkenler = _etken_adlari(etkin_maddeler, bitmis_testler)
    cikti = []

    gorunus = _bul(bitmis_testler, "görünüş")          # bitmiş/film görünüş
    ort_agirlik = _bul(bitmis_testler, "ortalama ağırlık")
    agirlik_tek = _bul(bitmis_testler, "ağırlık tekdüzeliği")
    dagilma = _bul(bitmis_testler, "dağılma")          # film için (30 dk)
    mikro = next((t for t in bitmis_testler if t.mikrobiyolojik), None)

    def mikro_kopya(op, yildiz):
        alt = mikro.alt_satirlar if mikro else [
            ("-Toplam Aerobik Mikroorganizma Sayısı", "≤10³ cfu/g"),
            ("-Küf ve Maya Sayısı", "≤10² cfu/g"), ("-E. coli", "0 cfu/g")]
        return _yeni("Mikrobiyolojik Kontrol", op, TabloTipi.MATRIS, mikro=True,
                     yildiz=yildiz, alt_satirlar=alt,
                     kaynak_spek=(mikro.spesifikasyon if mikro else None))

    def ilgili_bilesikler(op, yildiz):
        out = []
        for em in etkin_maddeler:
            if not em.impuriteler:
                continue
            bas = _yeni(f"{em.ad} İlgili Bileşikler", op, TabloTipi.IKI_NUMUNE, "", yildiz=yildiz)
            bas._grup_baslik = True
            out.append(bas)
            for imp in em.impuriteler:
                it = _yeni(f"—{imp.ad}", op, TabloTipi.IKI_NUMUNE, imp.limit_metni, yildiz=False)
                it.spesifikasyon.maksimum_deger = imp.maksimum_deger
                it._impurite = True
                out.append(it)
        return out

    def agirlik_kopya(op, tablet_alt=None):
        # tablet_alt verilirse (Tablet Baskı), alt satır spek'leri override edilir
        as1 = (agirlik_tek.aciklama_spek if agirlik_tek else "")
        as2 = (agirlik_tek.aciklama2_spek if agirlik_tek else "")
        if tablet_alt:
            as1 = tablet_alt.get("sapabilir", as1)
            as2 = tablet_alt.get("sapmamali", as2)
        return _yeni("Ağırlık Tekdüzeliği", op, TabloTipi.AGIRLIK_TEKDUZELIGI,
                     kaynak_spek=(agirlik_tek.spesifikasyon if agirlik_tek else None), ipk=True,
                     ac1=(agirlik_tek.aciklama_etiketi if agirlik_tek else "—20 tablette tek tek tabletlerden maksimum 2 tanesi bu limitten sapabilir."),
                     as1=as1,
                     ac2=(agirlik_tek.aciklama2_etiketi if agirlik_tek else "—Hiçbir tablet bu limitten sapmamalıdır."),
                     as2=as2)

    # ===================== KARIŞIM (Op 2) =====================
    if "Karıştırma" in operasyonlar:
        gor_kar = tablet_ipk.get("Görünüş_Karışım", "")
        if cift_katman:
            # Test bazında gruplu: önce tüm Görünüşler, sonra tüm Karışım Tek., vb.
            for em in etkenler:
                gor_em = tablet_ipk.get(f"Görünüş_Karışım::{em}", gor_kar)
                cikti.append(_yeni(f"{em} Görünüş", "Karıştırma", TabloTipi.TEK_SONUC,
                                   gor_em, ipk=True))
            for em in etkenler:
                cikti.append(_yeni(f"{em} Karışım Tekdüzeliği", "Karıştırma",
                                   TabloTipi.ON_NUMUNE, SABIT_KARISIM_SPEK, yildiz=True))
            for em in etkenler:
                cikti.append(_yeni(f"{em} Elek Testi", "Karıştırma", TabloTipi.TEK_SONUC, BILGI, yildiz=True))
            for em in etkenler:
                cikti.append(_yeni(f"{em} Bulk ve Tap Dansite", "Karıştırma", TabloTipi.TEK_SONUC, BILGI, yildiz=True))
        else:
            cikti.append(_yeni("Görünüş", "Karıştırma", TabloTipi.TEK_SONUC, gor_kar, ipk=True))
            cikti.append(_yeni("Elek Testi", "Karıştırma", TabloTipi.TEK_SONUC, BILGI, yildiz=True))
            cikti.append(_yeni("Bulk ve Tap Dansite", "Karıştırma", TabloTipi.TEK_SONUC, BILGI, yildiz=True))
            for em in etkenler:
                cikti.append(_yeni(f"{em} Karışım Tekdüzeliği", "Karıştırma",
                                   TabloTipi.ON_NUMUNE, SABIT_KARISIM_SPEK, yildiz=True))
        for em in etkenler:
            tes = _bul(bitmis_testler, "teşhis", etken=em)
            cikti.append(_yeni(f"{em} Teşhis", "Karıştırma", TabloTipi.TEK_SONUC,
                               kaynak_spek=(tes.spesifikasyon if tes else None)))
            mik = _bul(bitmis_testler, "miktar", etken=em)
            cikti.append(_yeni(f"{em} Miktar Tayini", "Karıştırma", TabloTipi.IKI_NUMUNE,
                               kaynak_spek=(mik.spesifikasyon if mik else None)))
        cikti += ilgili_bilesikler("Karıştırma", yildiz=True)
        # Çift katman: her etken için ayrı mikrobiyoloji (ikisi de *); tek katman: 1
        if cift_katman:
            for em in etkenler:
                mk = mikro_kopya("Karıştırma", yildiz=True)
                mk.ad = f"{em} Mikrobiyolojik Kontrol"
                cikti.append(mk)
        else:
            cikti.append(mikro_kopya("Karıştırma", yildiz=True))

    # ===================== TABLET BASKI (Op 3) =====================
    if "Tablet Baskı" in operasyonlar:
        cikti.append(_yeni("Görünüş", "Tablet Baskı", TabloTipi.TEK_SONUC,
                           tablet_ipk.get("Görünüş_Tablet", ""), ipk=True))
        cikti.append(_yeni("Ortalama Ağırlık", "Tablet Baskı", TabloTipi.BOS_NOKTA,
                           tablet_ipk.get("Ortalama Ağırlık_Tablet", ""), ipk=True))
        cikti.append(agirlik_kopya("Tablet Baskı", tablet_alt={
            "sapabilir": tablet_ipk.get("Ağırlık Tek Sapabilir", agirlik_tek.aciklama_spek if agirlik_tek else ""),
            "sapmamali": tablet_ipk.get("Ağırlık Tek Sapmamalı", agirlik_tek.aciklama2_spek if agirlik_tek else ""),
        }))
        cikti.append(_yeni("Kalınlık", "Tablet Baskı", TabloTipi.BOS_NOKTA,
                           tablet_ipk.get("Kalınlık", ""), ipk=True))
        cikti.append(_yeni("Çap", "Tablet Baskı", TabloTipi.BOS_NOKTA,
                           tablet_ipk.get("Çap", ""), ipk=True))
        cikti.append(_yeni("Sertlik", "Tablet Baskı", TabloTipi.TEK_SONUC,
                           tablet_ipk.get("Sertlik", ""), ipk=True))
        cikti.append(_yeni("Aşınma", "Tablet Baskı", TabloTipi.TEK_SONUC, ASINMA_SPEK, ipk=True))
        cikti.append(_yeni("Dağılma", "Tablet Baskı", TabloTipi.BOS_NOKTA, TABLET_DAGILMA, ipk=True))
        for em in etkenler:
            tes = _bul(bitmis_testler, "teşhis", etken=em)
            cikti.append(_yeni(f"{em} Teşhis", "Tablet Baskı", TabloTipi.TEK_SONUC,
                               kaynak_spek=(tes.spesifikasyon if tes else None)))
            mik = _bul(bitmis_testler, "miktar", etken=em)
            cikti.append(_yeni(f"{em} Miktar Tayini", "Tablet Baskı", TabloTipi.IKI_NUMUNE,
                               kaynak_spek=(mik.spesifikasyon if mik else None)))
            dis = _bul(bitmis_testler, "dissol", etken=em)
            if dis:
                cikti.append(_yeni(f"{em} Dissolüsyon (Q)", "Tablet Baskı", TabloTipi.BOS_NOKTA,
                                   kaynak_spek=dis.spesifikasyon))
        cikti += ilgili_bilesikler("Tablet Baskı", yildiz=True)
        cikti.append(mikro_kopya("Tablet Baskı", yildiz=True))

    # ===================== FİLM KAPLAMA (Op 4) =====================
    if "Film Kaplama" in operasyonlar:
        cikti.append(_yeni("Görünüş", "Film Kaplama", TabloTipi.TEK_SONUC,
                           kaynak_spek=(gorunus.spesifikasyon if gorunus else None), ipk=True))
        if ort_agirlik:
            cikti.append(_yeni("Ortalama Ağırlık", "Film Kaplama", TabloTipi.BOS_NOKTA,
                               kaynak_spek=ort_agirlik.spesifikasyon, ipk=True))
        cikti.append(agirlik_kopya("Film Kaplama"))
        if dagilma:
            cikti.append(_yeni("Dağılma", "Film Kaplama", TabloTipi.BOS_NOKTA,
                               kaynak_spek=dagilma.spesifikasyon, ipk=True))
        for em in etkenler:
            tes = _bul(bitmis_testler, "teşhis", etken=em)
            cikti.append(_yeni(f"{em} Teşhis", "Film Kaplama", TabloTipi.TEK_SONUC,
                               kaynak_spek=(tes.spesifikasyon if tes else None)))
            mik = _bul(bitmis_testler, "miktar", etken=em)
            cikti.append(_yeni(f"{em} Miktar Tayini", "Film Kaplama", TabloTipi.IKI_NUMUNE,
                               kaynak_spek=(mik.spesifikasyon if mik else None)))
            dis = _bul(bitmis_testler, "dissol", etken=em)
            if dis:
                cikti.append(_yeni(f"{em} Dissolüsyon (Q)", "Film Kaplama", TabloTipi.BOS_NOKTA,
                                   kaynak_spek=dis.spesifikasyon))
        cikti += ilgili_bilesikler("Film Kaplama", yildiz=False)  # film'de * YOK
        cikti.append(mikro_kopya("Film Kaplama", yildiz=True))

    # ===================== BLİSTERLEME (Op 5) =====================
    if "Blisterleme" in operasyonlar:
        cikti.append(_yeni("Sızdırmazlık", "Blisterleme", TabloTipi.TEK_SONUC,
                           SIZDIRMAZLIK_SPEK, ipk=True))
        cikti.append(mikro_kopya("Blisterleme", yildiz=False))  # blisterlemede * YOK

    # Tanımsız testleri (özel kuralla) ilgili aşamalara ekle
    _ozel_kuralli_ekle(cikti, bitmis_testler, ozel_test_kurallari)

    # Operasyon sırasına göre stable sort (aşama içi göreceli sıra korunur)
    op_sira = {"Karıştırma": 2, "Tablet Baskı": 3, "Film Kaplama": 4,
               "Dolum": 3, "Blisterleme": 5}
    cikti.sort(key=lambda t: op_sira.get(t.operasyon, 99))
    return cikti


def testleri_turet(bitmis_testler, operasyonlar, etkin_maddeler):
    return turet(bitmis_testler, etkin_maddeler, operasyonlar)
