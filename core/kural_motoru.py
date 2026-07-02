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
SABIT_ICERIK_SPEK = "Kabul Değeri (AV) ≤ L1 (L1=15,0)"
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
                  "dissol", "ilgili bilesik", "mikrobiyolojik", "sizdirmazlik",
                  "icerik tekduzeligi", "boyar madde", "enantiomerik",
                  "bulk", "tap dansite", "dansite"]
    # NOT: "nem" bilerek listede DEĞİL — Nem testinde program kullanıcıya hangi
    # aşamalarda uygulanacağını ve yıldızlı olup olmadığını sorar (özel kural).
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
    """Özel kuralı olan tanımsız testleri ilgili aşamalara dağıtır.
    Aynı (test adı + operasyon) çifti asla iki kez eklenmez (çift yazım önlenir)."""
    if not ozel_kurallar:
        return
    _eklenen = set()
    # Mevcut çıktıdaki (ad, op) çiftlerini de topla ki kuralla çift olmasın
    for t in cikti:
        _eklenen.add((_norm(t.ad), _norm(t.operasyon or "")))
    for t in bitmis_testler:
        kural = ozel_kurallar.get(t.ad)
        if not kural:
            continue
        spek = kural.get("spek") or (t.spesifikasyon.spesifikasyon_metni
                                     or t.spesifikasyon.metni_olustur())
        # asamalar listesini tekrarsız yap (sıra korunarak)
        asamalar = list(dict.fromkeys(kural.get("asamalar", [])))
        for op in asamalar:
            anahtar = (_norm(t.ad), _norm(op))
            if anahtar in _eklenen:
                continue  # bu test bu aşamada zaten var
            yildiz = op in kural.get("yildiz", [])
            cikti.append(_yeni(t.ad, op, TabloTipi.TEK_SONUC, spek,
                               ipk=kural.get("ipk", False), yildiz=yildiz))
            _eklenen.add(anahtar)


def turet(bitmis_testler, etkin_maddeler, operasyonlar,
          cift_katman=False, tablet_ipk=None, ozel_test_kurallari=None):
    tablet_ipk = tablet_ipk or {}
    etkenler = _etken_adlari(etkin_maddeler, bitmis_testler)
    cikti = []

    gorunus = _bul(bitmis_testler, "görünüş")          # bitmiş/film görünüş
    ort_agirlik = _bul(bitmis_testler, "ortalama ağırlık")
    agirlik_tek = (_bul(bitmis_testler, "ağırlık tekdüzeliği")
                   or _bul(bitmis_testler, "ağırlık sapması"))
    dagilma = _bul(bitmis_testler, "dağılma")          # film için (30 dk)
    mikro = next((t for t in bitmis_testler if t.mikrobiyolojik), None)

    def boyar_madde(op, boyar_test):
        """Boyar Madde Tanıması — İlgili Bileşikler ile aynı yapı: grup başlık +
        alt başlık(lar) (örn. Titanyum dioksit), her birinin kendi speki."""
        out = []
        bas = _yeni("Boyar Madde Tanıması", op, TabloTipi.TEK_SONUC,
                    kaynak_spek=boyar_test.spesifikasyon)
        bas._grup_baslik = True
        bas._boyar = True
        alt = list(getattr(boyar_test, "alt_satirlar", []) or [])
        bas._alt_basliklar = alt
        out.append(bas)
        for alt_ad, alt_spek in alt:
            it = _yeni(f"-{alt_ad}", op, TabloTipi.TEK_SONUC, alt_spek)
            it._boyar = True
            it._boyar_alt = True
            out.append(it)
        return out

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

    def enantiomerik_bilesikler(op, yildiz):
        """Enantiomerik İmpurite — İlgili Bileşikler ile AYNI yapı: grup başlık +
        alt satır(lar). Aşama başına BİR kez üretilir (çoğaltma yok)."""
        out = []
        for em in etkin_maddeler:
            if not getattr(em, "enantiomerik", None):
                continue
            bas = _yeni("Enantiomerik İmpurite", op, TabloTipi.IKI_NUMUNE, "", yildiz=yildiz)
            bas._grup_baslik = True
            bas._enantiomerik = True
            out.append(bas)
            for imp in em.enantiomerik:
                it = _yeni(f"—{imp.ad}", op, TabloTipi.IKI_NUMUNE, imp.limit_metni, yildiz=False)
                it.spesifikasyon.maksimum_deger = imp.maksimum_deger
                it._impurite = True
                it._enantiomerik = True
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
                cikti.append(_yeni(f"{em} Elek Testi", "Karıştırma", TabloTipi.TEK_SONUC, BILGI, ipk=True, yildiz=True))
            for em in etkenler:
                cikti.append(_yeni(f"{em} Bulk Dansite", "Karıştırma", TabloTipi.TEK_SONUC, BILGI, ipk=True, yildiz=True))
                cikti.append(_yeni(f"{em} Tap Dansite", "Karıştırma", TabloTipi.TEK_SONUC, BILGI, ipk=True, yildiz=True))
        else:
            cikti.append(_yeni("Görünüş", "Karıştırma", TabloTipi.TEK_SONUC, gor_kar, ipk=True))
            cikti.append(_yeni("Elek Testi", "Karıştırma", TabloTipi.TEK_SONUC, BILGI, ipk=True, yildiz=True))
            cikti.append(_yeni("Bulk Dansite", "Karıştırma", TabloTipi.TEK_SONUC, BILGI, ipk=True, yildiz=True))
            cikti.append(_yeni("Tap Dansite", "Karıştırma", TabloTipi.TEK_SONUC, BILGI, ipk=True, yildiz=True))
            for em in etkenler:
                cikti.append(_yeni(f"{em} Karışım Tekdüzeliği", "Karıştırma",
                                   TabloTipi.ON_NUMUNE, SABIT_KARISIM_SPEK, yildiz=True))
        for em in etkenler:
            tes = _bul(bitmis_testler, "teşhis", etken=em)
            cikti.append(_yeni(f"{em} Teşhis", "Karıştırma", TabloTipi.TEK_SONUC,
                               kaynak_spek=(tes.spesifikasyon if tes else None)))
            # NOT: Karışım aşamasında Miktar Tayini YOKTUR (kullanıcı kuralı).
            # Karışımda yalnızca Karışım Tekdüzeliği bulunur.
        cikti += ilgili_bilesikler("Karıştırma", yildiz=True)
        cikti += enantiomerik_bilesikler("Karıştırma", yildiz=True)
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
        _sertlik_bitmis = _bul(bitmis_testler, "sertlik")
        if tablet_ipk.get("Sertlik"):
            cikti.append(_yeni("Sertlik", "Tablet Baskı", TabloTipi.BOS_NOKTA,
                               tablet_ipk.get("Sertlik"), ipk=True))
        elif _sertlik_bitmis is not None:
            cikti.append(_yeni("Sertlik", "Tablet Baskı", TabloTipi.BOS_NOKTA,
                               kaynak_spek=_sertlik_bitmis.spesifikasyon, ipk=True))
        else:
            cikti.append(_yeni("Sertlik", "Tablet Baskı", TabloTipi.BOS_NOKTA, "", ipk=True))
        cikti.append(_yeni("Aşınma", "Tablet Baskı", TabloTipi.TEK_SONUC, ASINMA_SPEK, ipk=True))
        cikti.append(_yeni("Dağılma", "Tablet Baskı", TabloTipi.BOS_NOKTA, TABLET_DAGILMA, ipk=True))
        # Nem otomatik EKLENMEZ — kullanıcı hangi aşamalarda olacağını seçer
        # (özel kural mekanizması, _ozel_kuralli_ekle).
        for em in etkenler:
            tes = _bul(bitmis_testler, "teşhis", etken=em)
            cikti.append(_yeni(f"{em} Teşhis", "Tablet Baskı", TabloTipi.TEK_SONUC,
                               kaynak_spek=(tes.spesifikasyon if tes else None)))
            mik = _bul(bitmis_testler, "miktar", etken=em)
            cikti.append(_yeni(f"{em} Miktar Tayini", "Tablet Baskı", TabloTipi.IKI_NUMUNE,
                               kaynak_spek=(mik.spesifikasyon if mik else None)))
            # İçerik Tekdüzeliği HER ZAMAN eklenir (yıldızlı), Miktar Tayini'nden
            # hemen sonra. Bitmiş üründe tanımlıysa onun speki, yoksa AV varsayılanı.
            icerik = _bul(bitmis_testler, "içerik tekdüzeliği", etken=em)
            ic_spek = icerik.spesifikasyon if icerik else None
            it = _yeni(f"{em} İçerik Tekdüzeliği", "Tablet Baskı", TabloTipi.TEK_SONUC,
                       SABIT_ICERIK_SPEK if ic_spek is None else "",
                       kaynak_spek=ic_spek, yildiz=True)
            cikti.append(it)
            dis = _bul(bitmis_testler, "dissol", etken=em)
            if dis:
                cikti.append(_yeni(f"{em} Dissolüsyon", "Tablet Baskı", TabloTipi.BOS_NOKTA,
                                   kaynak_spek=dis.spesifikasyon))
        cikti += ilgili_bilesikler("Tablet Baskı", yildiz=True)
        cikti += enantiomerik_bilesikler("Tablet Baskı", yildiz=True)
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
        # Nem otomatik EKLENMEZ — kullanıcı seçer (özel kural).
        for em in etkenler:
            tes = _bul(bitmis_testler, "teşhis", etken=em)
            cikti.append(_yeni(f"{em} Teşhis", "Film Kaplama", TabloTipi.TEK_SONUC,
                               kaynak_spek=(tes.spesifikasyon if tes else None)))
            mik = _bul(bitmis_testler, "miktar", etken=em)
            cikti.append(_yeni(f"{em} Miktar Tayini", "Film Kaplama", TabloTipi.IKI_NUMUNE,
                               kaynak_spek=(mik.spesifikasyon if mik else None)))
            # İçerik Tekdüzeliği HER ZAMAN (yıldızlı), Miktar Tayini'nden sonra
            icerik = _bul(bitmis_testler, "içerik tekdüzeliği", etken=em)
            ic_spek = icerik.spesifikasyon if icerik else None
            cikti.append(_yeni(f"{em} İçerik Tekdüzeliği", "Film Kaplama", TabloTipi.TEK_SONUC,
                               SABIT_ICERIK_SPEK if ic_spek is None else "",
                               kaynak_spek=ic_spek, yildiz=True))
            dis = _bul(bitmis_testler, "dissol", etken=em)
            if dis:
                cikti.append(_yeni(f"{em} Dissolüsyon", "Film Kaplama", TabloTipi.BOS_NOKTA,
                                   kaynak_spek=dis.spesifikasyon))
        boyar = _bul(bitmis_testler, "boyar madde")
        if boyar:
            cikti += boyar_madde("Film Kaplama", boyar)
        cikti += ilgili_bilesikler("Film Kaplama", yildiz=False)  # film'de * YOK
        cikti += enantiomerik_bilesikler("Film Kaplama", yildiz=False)
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

    # Her aşamada kimyasal test sırası düzeltmesi:
    #   ... Dissolüsyon → (Nem, Boyar Madde) → İlgili Bileşikler(+alt) →
    #   Enantiomerik(+alt) → Mikrobiyolojik(son).
    # İPK testleri ve grup başlığı+alt satır blokları bozulmadan taşınır.
    def _mikro_mu(t):
        return getattr(t, "mikrobiyolojik", False) or "mikrobiyolojik" in _norm(t.ad)

    def _grup_anahtar(t):
        """Bir testi kimyasal sıralama grubuna eşler (İPK ise None → yerinde kalır)."""
        if getattr(t, "ipk", False):
            return None
        a = _norm(t.ad)
        if _mikro_mu(t):
            return 9
        if getattr(t, "_enantiomerik", False) or "enantiomerik" in a:
            return 8
        if getattr(t, "_impurite", False) or "ilgili bilesik" in a:
            return 7
        if getattr(t, "_boyar", False) or "boyar madde" in a:
            return 6
        if "nem" in a:
            return 5
        if "dissol" in a:
            return 4
        if "icerik tekduzeligi" in a:
            return 3
        if "miktar tayini" in a:
            return 2
        if "teshis" in a or "karisim tekduzeligi" in a:
            return 1
        return None  # tanınmayan kimyasal → yerinde kalır

    yeniden = []
    i = 0
    n = len(cikti)
    while i < n:
        op = cikti[i].operasyon
        grup = []
        while i < n and cikti[i].operasyon == op:
            grup.append(cikti[i]); i += 1
        # Kimyasal testleri sabit sıraya diz; İPK ve grup-başlık+alt bloklarını koru.
        # Blok oluştur: grup başlığı + ardışık alt satırlar bir arada.
        bloklar = []
        j = 0
        while j < len(grup):
            t = grup[j]
            blok = [t]
            if getattr(t, "_grup_baslik", False):
                k = j + 1
                while k < len(grup) and (getattr(grup[k], "_impurite", False)
                                          or getattr(grup[k], "_boyar_alt", False)):
                    blok.append(grup[k]); k += 1
                j = k
            else:
                j += 1
            bloklar.append(blok)
        # Sıralama anahtarı: İPK (None) olanlar yerinde; kimyasal olanlar gruba göre
        def _blok_sira(idx_blok):
            idx, blok = idx_blok
            ga = _grup_anahtar(blok[0])
            return (1, ga, idx) if ga is not None else (0, 0, idx)
        sirali = [b for _, b in sorted(enumerate(bloklar), key=_blok_sira)]
        for blok in sirali:
            yeniden.extend(blok)
    return yeniden


def testleri_turet(bitmis_testler, operasyonlar, etkin_maddeler):
    return turet(bitmis_testler, etkin_maddeler, operasyonlar)
